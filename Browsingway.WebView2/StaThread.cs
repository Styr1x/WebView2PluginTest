using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using TerraFX.Interop.Windows;
using static TerraFX.Interop.Windows.Windows;

namespace Browsingway.WebView2;

/// <summary>
/// Manages a dedicated STA (Single-Threaded Apartment) thread for WebView2 operations.
/// WebView2 requires STA COM threading, and this helper provides methods to marshal
/// work onto the STA thread safely.
/// </summary>
internal sealed class StaThread : IDisposable, IAsyncDisposable
{
    private Thread? _staThread;
    private BlockingCollection<Action>? _workQueue;
    private TaskCompletionSource<bool>? _threadReadyTcs;
    private volatile bool _disposed;
    private IntPtr _dispatcherQueueController;

    /// <summary>
    /// Gets whether the STA thread has been started and is ready.
    /// </summary>
    public bool IsRunning => _staThread != null && !_disposed;

    /// <summary>
    /// Fired when the STA thread is ready. Handlers run on the STA thread
    /// and can perform initialization that requires STA context.
    /// </summary>
    public event Action? ThreadReady;

    /// <summary>
    /// Starts the STA thread and waits for it to be ready.
    /// </summary>
    public async Task StartAsync()
    {
        if (_staThread != null)
            throw new InvalidOperationException("STA thread already started");

        _workQueue = new BlockingCollection<Action>();
        _threadReadyTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(StaThreadProc)
        {
            Name = "WebView2 STA Thread",
            IsBackground = true
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        await _threadReadyTcs.Task;
    }

    /// <summary>
    /// Runs an action on the STA thread and waits for completion.
    /// </summary>
    public void Run(Action action)
    {
        if (_workQueue == null || _staThread == null)
            throw new InvalidOperationException("STA thread not initialized");

        if (_disposed)
            throw new ObjectDisposedException(nameof(StaThread));

        if (Thread.CurrentThread == _staThread)
        {
            // Already on STA thread, execute directly
            action();
            return;
        }

        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        _workQueue.Add(() =>
        {
            try
            {
                action();
                tcs.SetResult(true);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        tcs.Task.GetAwaiter().GetResult();
    }

    /// <summary>
    /// Runs an async function on the STA thread and returns the result.
    /// </summary>
    public unsafe Task<T> RunAsync<T>(Func<Task<T>> asyncFunc)
    {
        if (_workQueue == null || _staThread == null)
            throw new InvalidOperationException("STA thread not initialized");

        if (_disposed)
            throw new ObjectDisposedException(nameof(StaThread));

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _workQueue.Add(() =>
        {
            try
            {
                // Run the async work and pump messages while waiting
                var task = asyncFunc();
                while (!task.IsCompleted)
                {
                    // Pump COM/Win32 messages to keep WebView2 responsive
                    MSG msg;
                    while (PeekMessageW(&msg, HWND.NULL, 0, 0, PM.PM_REMOVE))
                    {
                        TranslateMessage(&msg);
                        DispatchMessageW(&msg);
                    }
                    Thread.Sleep(1);
                }

                if (task.IsFaulted)
                    tcs.SetException(task.Exception!.InnerExceptions);
                else if (task.IsCanceled)
                    tcs.SetCanceled();
                else
                    tcs.SetResult(task.Result);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        return tcs.Task;
    }

    /// <summary>
    /// Runs an async action on the STA thread.
    /// </summary>
    public Task RunAsync(Func<Task> asyncAction)
    {
        return RunAsync(async () =>
        {
            await asyncAction();
            return true;
        });
    }

    /// <summary>
    /// Posts an action to the STA thread without waiting (fire-and-forget).
    /// </summary>
    public void Post(Action action)
    {
        if (_workQueue == null)
            throw new InvalidOperationException("STA thread not initialized");

        if (_disposed)
            return; // Silently ignore posts after disposal

        _workQueue.Add(action);
    }

    private unsafe void StaThreadProc()
    {
        try
        {
            // Initialize COM as STA
            CoInitializeEx(null, (uint)COINIT.COINIT_APARTMENTTHREADED);

            try
            {
                // Create dispatcher queue for WinRT composition on this STA thread
                CreateDispatcherQueue();

                // Fire thread ready event for initialization
                ThreadReady?.Invoke();

                // Signal that thread is ready
                _threadReadyTcs?.SetResult(true);

                // Message pump loop
                while (!_disposed)
                {
                    // Process work items
                    if (_workQueue!.TryTake(out var work, 10))
                    {
                        try
                        {
                            work();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"STA work item exception: {ex}");
                        }
                    }

                    // Pump Win32 messages
                    MSG msg;
                    while (PeekMessageW(&msg, HWND.NULL, 0, 0, PM.PM_REMOVE))
                    {
                        if (msg.message == WM.WM_QUIT)
                        {
                            return;
                        }
                        TranslateMessage(&msg);
                        DispatchMessageW(&msg);
                    }
                }
            }
            finally
            {
                // Release the dispatcher queue controller before uninitializing COM
                ReleaseDispatcherQueue();
                CoUninitialize();
            }
        }
        catch (Exception ex)
        {
            _threadReadyTcs?.TrySetException(ex);
        }
    }

    // Dispatcher queue creation (coremessaging.dll) - not in TerraFX
    private const int DQTYPE_THREAD_CURRENT = 2;
    private const int DQTAT_COM_STA = 2;

    [StructLayout(LayoutKind.Sequential)]
    private struct DispatcherQueueOptions
    {
        public int dwSize;
        public int threadType;
        public int apartmentType;
    }

    [DllImport("coremessaging.dll", EntryPoint = "CreateDispatcherQueueController")]
    private static extern int CreateDispatcherQueueController(
        DispatcherQueueOptions options,
        out IntPtr dispatcherQueueController);

    private void CreateDispatcherQueue()
    {
        var options = new DispatcherQueueOptions
        {
            dwSize = Marshal.SizeOf<DispatcherQueueOptions>(),
            threadType = DQTYPE_THREAD_CURRENT,
            apartmentType = DQTAT_COM_STA
        };

        // Store the controller so we can release it on disposal
        var hr = CreateDispatcherQueueController(options, out _dispatcherQueueController);
        if (hr < 0 && _dispatcherQueueController == IntPtr.Zero)
        {
            System.Diagnostics.Debug.WriteLine($"CreateDispatcherQueueController failed: 0x{hr:X8}");
        }
    }

    private void ReleaseDispatcherQueue()
    {
        if (_dispatcherQueueController != IntPtr.Zero)
        {
            Marshal.Release(_dispatcherQueueController);
            _dispatcherQueueController = IntPtr.Zero;
        }
    }

    /// <summary>
    /// Disposes the helper asynchronously, waiting for the STA thread to exit.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        if (_staThread != null && _workQueue != null)
        {
            _workQueue.CompleteAdding();
            await Task.Run(() => _staThread.Join(TimeSpan.FromSeconds(5)));
            _workQueue.Dispose();
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes the helper synchronously.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        if (_staThread != null && _workQueue != null)
        {
            _workQueue.CompleteAdding();
            _staThread.Join(TimeSpan.FromSeconds(5));
            _workQueue.Dispose();
        }

        GC.SuppressFinalize(this);
    }
}
