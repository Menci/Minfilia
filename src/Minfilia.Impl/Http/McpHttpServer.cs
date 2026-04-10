using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Minfilia.Http;

internal sealed class McpHttpServer(McpServerOptions _serverOptions) : IDisposable
{
    private HttpListener? _listener;
    private CancellationTokenSource? _cts;

    private static readonly JsonTypeInfo<JsonRpcMessage> MessageTypeInfo =
        (JsonTypeInfo<JsonRpcMessage>)McpJsonUtilities.DefaultOptions.GetTypeInfo(typeof(JsonRpcMessage));

    private static readonly JsonTypeInfo<JsonRpcError> ErrorTypeInfo =
        (JsonTypeInfo<JsonRpcError>)McpJsonUtilities.DefaultOptions.GetTypeInfo(typeof(JsonRpcError));

    public async Task RunAsync(string prefix, CancellationToken cancellationToken = default)
    {
        _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _listener = new HttpListener();
        _listener.Prefixes.Add(prefix);
        _listener.Start();

        Console.WriteLine($"MCP server listening on {prefix}");

        while (!_cts.Token.IsCancellationRequested)
        {
            try
            {
                var ctx = await _listener.GetContextAsync().ConfigureAwait(false);
                _ = HandleRequestAsync(ctx, _cts.Token);
            }
            catch (HttpListenerException) when (_cts.Token.IsCancellationRequested)
            {
                break;
            }
            catch (ObjectDisposedException)
            {
                break;
            }
        }
    }

    private async Task HandleRequestAsync(HttpListenerContext ctx, CancellationToken cancellationToken)
    {
        try
        {
            switch (ctx.Request.HttpMethod)
            {
                case "POST":
                    await HandlePostAsync(ctx, cancellationToken).ConfigureAwait(false);
                    break;
                case "DELETE":
                    ctx.Response.StatusCode = 200;
                    ctx.Response.Close();
                    break;
                default:
                    ctx.Response.StatusCode = 405;
                    ctx.Response.Close();
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error handling request: {ex.Message}");
            try
            {
                ctx.Response.StatusCode = 500;
                ctx.Response.Close();
            }
            catch { }
        }
    }

    private async Task HandlePostAsync(HttpListenerContext ctx, CancellationToken cancellationToken)
    {
        JsonRpcMessage? message;
        try
        {
            using var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8);
            var json = await reader.ReadToEndAsync().ConfigureAwait(false);
            message = JsonSerializer.Deserialize(json, MessageTypeInfo);
        }
        catch
        {
            await WriteErrorAsync(ctx, "Bad Request: Invalid JSON-RPC message.", 400).ConfigureAwait(false);
            return;
        }

        if (message == null)
        {
            await WriteErrorAsync(ctx, "Bad Request: Empty message.", 400).ConfigureAwait(false);
            return;
        }

        var transport = new StreamableHttpServerTransport { Stateless = true };
        var server = McpServer.Create(transport, _serverOptions);

        using var requestCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var runTask = server.RunAsync(requestCts.Token);

        ctx.Response.ContentType = "text/event-stream; charset=utf-8";
        ctx.Response.ContentEncoding = Encoding.UTF8;
        ctx.Response.Headers["Cache-Control"] = "no-cache,no-store";
        ctx.Response.Headers["X-Accel-Buffering"] = "no";

        try
        {
            var wrote = await transport.HandlePostRequestAsync(message, ctx.Response.OutputStream, cancellationToken).ConfigureAwait(false);
            if (!wrote)
            {
                ctx.Response.StatusCode = 202;
            }
        }
        finally
        {
            await transport.DisposeAsync().ConfigureAwait(false);
            requestCts.Cancel();

            try { await runTask.ConfigureAwait(false); }
            catch (OperationCanceledException) { }

            ctx.Response.Close();
        }
    }

    private static async Task WriteErrorAsync(HttpListenerContext ctx, string errorMessage, int statusCode)
    {
        var error = new JsonRpcError
        {
            Error = new JsonRpcErrorDetail
            {
                Code = -32000,
                Message = errorMessage,
            },
        };

        ctx.Response.StatusCode = statusCode;
        ctx.Response.ContentType = "application/json; charset=utf-8";
        ctx.Response.ContentEncoding = Encoding.UTF8;
        var json = JsonSerializer.Serialize(error, ErrorTypeInfo);
        var bytes = Encoding.UTF8.GetBytes(json);
        await ctx.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        ctx.Response.Close();
    }

    public void Dispose()
    {
        _cts?.Cancel();
        _listener?.Stop();
        _listener?.Close();
        _cts?.Dispose();
    }
}
