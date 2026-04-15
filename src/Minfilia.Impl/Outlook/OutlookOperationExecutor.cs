using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;

namespace Minfilia.Outlook;

internal sealed class OutlookOperationExecutor
{
    private readonly SemaphoreSlim _gate = new(1, 1);

    public async Task ExecuteAsync(Action<dynamic> action)
    {
        await _gate.WaitAsync();

        try
        {
            using var session = new OutlookSession();
            await WaitForReadyAsync(session);
            await session.ExecuteAsync(action);
        }
        catch (Exception ex)
        {
            throw WrapOperationException(ex);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<T> ExecuteAsync<T>(Func<dynamic, T> func)
    {
        await _gate.WaitAsync();

        try
        {
            using var session = new OutlookSession();
            await WaitForReadyAsync(session);
            return await session.ExecuteAsync(func);
        }
        catch (Exception ex)
        {
            throw WrapOperationException(ex);
        }
        finally
        {
            _gate.Release();
        }
    }

    private static async Task WaitForReadyAsync(OutlookSession session)
    {
        try
        {
            await session.WaitForReadyAsync();
        }
        catch (Exception ex)
        {
            throw WrapInitializationException(ex);
        }
    }

    private static Exception WrapInitializationException(Exception ex)
    {
        if (ex is McpException)
            return ex;

        if (ex is COMException comException)
            return new McpException($"Failed to initialize Outlook COM session (HRESULT 0x{unchecked((uint)comException.HResult):X8}): {comException.Message}");

        return new McpException($"Failed to initialize Outlook COM session: {ex.Message}");
    }

    private static Exception WrapOperationException(Exception ex)
    {
        if (ex is McpException)
            return ex;

        if (ex is COMException comException)
            return new McpException($"Outlook COM call failed (HRESULT 0x{unchecked((uint)comException.HResult):X8}): {comException.Message}");

        return ex;
    }
}
