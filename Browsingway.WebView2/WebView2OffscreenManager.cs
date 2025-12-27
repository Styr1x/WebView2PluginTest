using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using TerraFX.Interop.DirectX;
using Windows.UI.Composition;

namespace Browsingway.WebView2;

/// <summary>
/// Manages multiple WebView2 offscreen views. Accepts an external D3D11 device
/// and creates shared resources for all views.
/// </summary>
public sealed class WebView2OffscreenManager : IDisposable, IAsyncDisposable
{
    private readonly unsafe ID3D11Device* _device;
    private readonly List<WebView2OffscreenView> _views = [];
    private readonly Lock _lock = new();

    private StaThread? _staThread;

    private CoreWebView2Environment? _environment;
    private Compositor? _compositor;
    private bool _disposed;
    private bool _initialized;

    /// <summary>
    /// Gets the currently focused view, or null if no view is focused.
    /// </summary>
    public WebView2OffscreenView? FocusedView { get; private set; }

    /// <summary>
    /// Fired when the cursor of the focused view changes.
    /// The IntPtr is the cursor handle (HCURSOR).
    /// </summary>
    public event EventHandler<IntPtr>? CursorChanged;

    /// <summary>
    /// Creates a new WebView2OffscreenManager with the specified D3D11 device.
    /// </summary>
    /// <param name="device">The D3D11 device to use for all texture operations. Caller retains ownership.</param>
    public unsafe WebView2OffscreenManager(ID3D11Device* device)
    {
        if (device == null)
            throw new ArgumentNullException(nameof(device));

        _device = device;
    }

    /// <summary>
    /// Initializes shared resources. Must be called before creating views.
    /// This starts a dedicated STA thread for WebView2 operations.
    /// </summary>
    public async Task InitializeAsync(string userDataFolder)
    {
        if (_initialized)
            return;

        // Create and start the STA thread helper
        _staThread = new StaThread();
        await _staThread.StartAsync();

        // Create shared WebView2 environment on STA thread
        await _staThread.RunAsync(async () =>
        {
            var envOptions = new CoreWebView2EnvironmentOptions();
            _environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

            // Create shared compositor
            _compositor = new Compositor();
        });

        _initialized = true;
    }

    /// <summary>
    /// Creates a new WebView2 offscreen view with the specified dimensions.
    /// </summary>
    /// <param name="width">Initial width in pixels.</param>
    /// <param name="height">Initial height in pixels.</param>
    /// <returns>The created view. Call InitializeAsync() on the view before use.</returns>
    public unsafe WebView2OffscreenView CreateView(int width, int height)
    {
        if (!_initialized)
            throw new InvalidOperationException("Manager must be initialized before creating views. Call InitializeAsync() first.");

        if (width <= 0)
            throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0)
            throw new ArgumentOutOfRangeException(nameof(height));

        var view = new WebView2OffscreenView(
            this,
            _staThread!,
            _device,
            _environment!,
            _compositor!,
            width,
            height);

        lock (_lock)
        {
            _views.Add(view);
        }

        return view;
    }

    public void ClearFocus()
    {
        SetFocusedView(null);
    }

    /// <summary>
    /// Removes a view from management. Called internally when view is disposed.
    /// </summary>
    internal void RemoveView(WebView2OffscreenView view)
    {
        lock (_lock)
        {
            _views.Remove(view);
            if (FocusedView == view)
            {
                FocusedView = null;
            }
        }
    }

    /// <summary>
    /// Sets focus to the specified view. Called internally by view.SetFocus().
    /// </summary>
    internal void SetFocusedView(WebView2OffscreenView? view)
    {
        WebView2OffscreenView? oldFocused;

        lock (_lock)
        {
            if (FocusedView == view)
                return;

            oldFocused = FocusedView;
            FocusedView = view;
        }

        // Notify old view it lost focus
        oldFocused?.OnFocusLost();

        // Notify new view it gained focus
        view?.OnFocusGained();

        // Fire cursor changed event with new view's cursor
        if (view != null)
        {
            CursorChanged?.Invoke(this, view.Cursor);
        }
    }

    /// <summary>
    /// Called by view when its cursor changes. Fires CursorChanged if view is focused.
    /// </summary>
    internal void OnViewCursorChanged(WebView2OffscreenView view, IntPtr cursor)
    {
        if (FocusedView == view)
        {
            CursorChanged?.Invoke(this, cursor);
        }
    }

    #region Input Routing

    /// <summary>
    /// Sends a mouse move event to the focused view.
    /// </summary>
    /// <param name="x">X coordinate relative to view.</param>
    /// <param name="y">Y coordinate relative to view.</param>
    /// <param name="modifiers">Current modifier key and mouse button state.</param>
    public void SendMouseMove(int x, int y, VirtualKeys modifiers = VirtualKeys.None)
    {
        FocusedView?.HandleMouseMove(x, y, modifiers);
    }

    /// <summary>
    /// Sends a mouse button event to the focused view.
    /// </summary>
    /// <param name="button">The mouse button.</param>
    /// <param name="down">True if button is pressed, false if released.</param>
    /// <param name="x">X coordinate relative to view.</param>
    /// <param name="y">Y coordinate relative to view.</param>
    /// <param name="modifiers">Current modifier key and mouse button state.</param>
    public void SendMouseButton(MouseButton button, bool down, int x, int y, VirtualKeys modifiers = VirtualKeys.None)
    {
        FocusedView?.HandleMouseButton(button, down, x, y, modifiers);
    }

    /// <summary>
    /// Sends a mouse wheel event to the focused view.
    /// </summary>
    /// <param name="delta">Wheel delta (positive = scroll up/right, negative = scroll down/left).</param>
    /// <param name="x">X coordinate relative to view.</param>
    /// <param name="y">Y coordinate relative to view.</param>
    /// <param name="horizontal">True for horizontal scroll, false for vertical.</param>
    /// <param name="modifiers">Current modifier key and mouse button state.</param>
    public void SendMouseWheel(int delta, int x, int y, bool horizontal = false, VirtualKeys modifiers = VirtualKeys.None)
    {
        FocusedView?.HandleMouseWheel(delta, x, y, horizontal, modifiers);
    }

    /// <summary>
    /// Sends a mouse leave event to the focused view.
    /// </summary>
    public void SendMouseLeave()
    {
        FocusedView?.HandleMouseLeave();
    }

    /// <summary>
    /// Sends a keyboard message to the focused view.
    /// </summary>
    /// <param name="msg">The Windows message (WM_KEYDOWN, WM_KEYUP, WM_CHAR, etc.).</param>
    /// <param name="wParam">The wParam of the message.</param>
    /// <param name="lParam">The lParam of the message.</param>
    public void SendKeyboardMessage(uint msg, IntPtr wParam, IntPtr lParam)
    {
        FocusedView?.HandleKeyboardMessage(msg, wParam, lParam);
    }

    #endregion

    /// <summary>
    /// Disposes the manager asynchronously, ensuring proper cleanup of all views and the STA thread.
    /// This is the preferred disposal method.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Get all views and clear the list
        WebView2OffscreenView[] viewsCopy;
        lock (_lock)
        {
            viewsCopy = _views.ToArray();
            _views.Clear();
            FocusedView = null;
        }

        if (_staThread != null)
        {
            // Dispose all views asynchronously on STA thread
            foreach (var view in viewsCopy)
            {
                try
                {
                    await _staThread.RunAsync(() =>
                    {
                        view.Dispose();
                        return Task.CompletedTask;
                    });
                }
                catch
                {
                    // Thread may already be shutting down
                }
            }

            // Dispose shared resources on STA thread
            try
            {
                await _staThread.RunAsync(() =>
                {
                    _compositor?.Dispose();
                    return Task.CompletedTask;
                });
            }
            catch
            {
                // Thread may already be shutting down
            }

            // Dispose the STA thread helper
            await _staThread.DisposeAsync();
        }
        else
        {
            // Fallback if thread wasn't started
            foreach (var view in viewsCopy)
            {
                view.Dispose();
            }

            _compositor?.Dispose();
        }

        GC.SuppressFinalize(this);

        // Note: We don't dispose _environment as it's managed by WinRT
        // Note: We don't release _device as caller owns it
    }

    /// <summary>
    /// Disposes the manager synchronously. Prefer DisposeAsync for proper cleanup.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Dispose all views on STA thread
        WebView2OffscreenView[] viewsCopy;
        lock (_lock)
        {
            viewsCopy = _views.ToArray();
            _views.Clear();
            FocusedView = null;
        }

        if (_staThread != null)
        {
            // Dispose views on STA thread
            foreach (var view in viewsCopy)
            {
                try
                {
                    _staThread.Run(() => view.Dispose());
                }
                catch
                {
                    // Thread may already be shutting down
                }
            }

            // Dispose shared resources on STA thread
            try
            {
                _staThread.Run(() => { _compositor?.Dispose(); });
            }
            catch
            {
                // Thread may already be shutting down
            }

            // Dispose the STA thread helper
            _staThread.Dispose();
        }
        else
        {
            // Fallback if thread wasn't started
            foreach (var view in viewsCopy)
            {
                view.Dispose();
            }

            _compositor?.Dispose();
        }

        GC.SuppressFinalize(this);

        // Note: We don't dispose _environment as it's managed by WinRT
        // Note: We don't release _device as caller owns it
    }
}
