using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Minfilia.Outlook;

internal sealed class OutlookSession : IDisposable
{
    private readonly Thread _staThread;
    private readonly BlockingCollection<Action> _queue = [];
    private readonly TaskCompletionSource<bool> _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);

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
        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        Enqueue(() =>
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
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        Enqueue(() =>
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

    private void Enqueue(Action action)
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(OutlookSession));

        try
        {
            _queue.Add(action);
        }
        catch (InvalidOperationException) when (_disposed)
        {
            throw new ObjectDisposedException(nameof(OutlookSession));
        }
    }

    private void StaThreadLoop()
    {
        try
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

            foreach (var action in _queue.GetConsumingEnumerable())
            {
                action();
            }
        }
        finally
        {
            ReleaseComObjects();
        }
    }

    private void ReleaseComObjects()
    {
        var namespaceObject = (object?)_namespace;
        var outlookObject = (object?)_outlook;

        _namespace = null;
        _outlook = null;

        ReleaseComObject(namespaceObject, "namespace");
        ReleaseComObject(outlookObject, "application");
    }

    private static void ReleaseComObject(object? comObject, string name)
    {
        if (comObject == null)
            return;

        try
        {
            if (Marshal.IsComObject(comObject))
                Marshal.FinalReleaseComObject(comObject);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: failed to release Outlook {name} COM object: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _queue.CompleteAdding();
        if (!_staThread.Join(TimeSpan.FromSeconds(5)))
        {
            Console.Error.WriteLine("Warning: Outlook STA thread did not stop within 5 seconds.");
            return;
        }

        _queue.Dispose();
    }
}
