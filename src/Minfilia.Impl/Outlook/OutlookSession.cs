using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace Minfilia.Outlook;

internal sealed class OutlookSession : IDisposable
{
    private readonly Thread _staThread;
    private readonly BlockingCollection<Action> _queue = [];
    private readonly TaskCompletionSource<bool> _ready = new();
    private readonly CancellationTokenSource _cts = new();

    private dynamic? _outlook;
    private dynamic? _namespace;
    private bool _disposed;

    public OutlookSession()
    {
        _staThread = new Thread(StaThreadLoop) { IsBackground = true };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
    }

    public Task WaitForReadyAsync() => _ready.Task;

    public Task ExecuteAsync(Action<dynamic> action)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(OutlookSession));

        var tcs = new TaskCompletionSource<bool>();
        _queue.Add(() =>
        {
            try
            {
                action(_namespace!);
                tcs.SetResult(true);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        return tcs.Task;
    }

    public Task<T> ExecuteAsync<T>(Func<dynamic, T> func)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(OutlookSession));

        var tcs = new TaskCompletionSource<T>();
        _queue.Add(() =>
        {
            try
            {
                var result = func(_namespace!);
                tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        return tcs.Task;
    }

    private void StaThreadLoop()
    {
        try
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null)
            {
                _ready.SetException(new InvalidOperationException(
                    "Outlook.Application COM class not found. Is Outlook installed?"));
                return;
            }

            _outlook = Activator.CreateInstance(type);
            _namespace = _outlook!.GetNamespace("MAPI");
            _namespace!.Logon();
            _ready.SetResult(true);
        }
        catch (Exception ex)
        {
            _ready.SetException(ex);
            return;
        }

        try
        {
            foreach (var action in _queue.GetConsumingEnumerable(_cts.Token))
            {
                action();
            }
        }
        catch (OperationCanceledException)
        {
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _cts.Cancel();
        _queue.CompleteAdding();
        _staThread.Join(TimeSpan.FromSeconds(5));

        _queue.Dispose();
        _cts.Dispose();
    }
}
