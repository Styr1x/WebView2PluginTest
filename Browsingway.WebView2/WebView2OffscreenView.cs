using System;
using System.Drawing;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using TerraFX.Interop.DirectX;
using TerraFX.Interop.Windows;
using Windows.UI.Composition;
using static TerraFX.Interop.Windows.Windows;

namespace Browsingway.WebView2;

/// <summary>
/// Represents a single WebView2 offscreen view that renders to a texture.
/// </summary>
public sealed class WebView2OffscreenView : IDisposable, IAsyncDisposable
{
    private readonly WebView2OffscreenManager _manager;
    private readonly StaThread _staThread;
    private readonly unsafe ID3D11Device* _device;
    private readonly CoreWebView2Environment _environment;
    private readonly Compositor _compositor;

    private HiddenWindow? _hiddenWindow;

    private int _width;
    private int _height;

    private CoreWebView2CompositionController? _compositionController;
    private CoreWebView2Controller? _controller;
    private CoreWebView2? _webView;
    private ContainerVisual? _rootVisual;
    private ContainerVisual? _webViewVisual;
    private GraphicsCaptureHelper? _captureHelper;

    private IntPtr _cursor;
    private HWND _webViewHwnd; // Internal WebView2 window handle for keyboard input
    private bool _isFocused;
    private bool _disposed;
    private bool _initialized;
    private byte _clickThroughAlphaThreshold; // Alpha threshold for click-through (0-255, 0 = disabled)

    /// <summary>
    /// Gets the current width of the view.
    /// </summary>
    public int Width => _width;

    /// <summary>
    /// Gets the current height of the view.
    /// </summary>
    public int Height => _height;

    /// <summary>
    /// Gets whether this view currently has focus.
    /// </summary>
    public bool IsFocused => _isFocused;

    /// <summary>
    /// Gets the current URL being displayed.
    /// </summary>
    public string? CurrentUrl => _webView?.Source;

    /// <summary>
    /// Gets the current cursor handle for this view.
    /// </summary>
    public IntPtr Cursor => _cursor;

    /// <summary>
    /// Gets or sets the alpha threshold for click-through behavior.
    /// Pixels with alpha values less than or equal to this threshold will be "clicked through".
    /// Set to 0 to disable click-through (all pixels are clickable).
    /// Set to 255 to make fully transparent pixels click-through only.
    /// </summary>
    public byte ClickThroughAlphaThreshold
    {
        get => _clickThroughAlphaThreshold;
        set => _clickThroughAlphaThreshold = value;
    }

    /// <summary>
    /// Gets the alpha value at the specified position in the view.
    /// Returns 255 (fully opaque) if position is out of bounds or texture is not available.
    /// </summary>
    /// <param name="x">X coordinate relative to view.</param>
    /// <param name="y">Y coordinate relative to view.</param>
    /// <returns>Alpha value (0-255) at the specified position.</returns>
    public byte GetAlphaAtPosition(int x, int y)
    {
        if (_captureHelper == null || x < 0 || y < 0 || x >= _width || y >= _height)
            return 255; // Default to opaque if out of bounds

        return _captureHelper.GetAlphaAtPosition(x, y);
    }

    /// <summary>
    /// Checks if a click at the specified position should pass through this view.
    /// </summary>
    /// <param name="x">X coordinate relative to view.</param>
    /// <param name="y">Y coordinate relative to view.</param>
    /// <returns>True if the click should pass through (not be handled by this view).</returns>
    public bool ShouldClickThrough(int x, int y)
    {
        if (_clickThroughAlphaThreshold == 0)
            return false; // Click-through disabled

        byte alpha = GetAlphaAtPosition(x, y);
        return alpha <= _clickThroughAlphaThreshold;
    }

    /// <summary>
    /// Fired when the view is ready to use.
    /// </summary>
    public event EventHandler? Ready;

    /// <summary>
    /// Fired when navigation completes. The string parameter is the URL.
    /// </summary>
    public event EventHandler<string>? NavigationCompleted;

    /// <summary>
    /// Fired when a new frame has been captured.
    /// The IntPtr parameter is the ID3D11Texture2D* pointer.
    /// </summary>
    public event EventHandler<IntPtr>? FrameCaptured;

    /// <summary>
    /// Fired when the cursor changes.
    /// The IntPtr parameter is the HCURSOR handle.
    /// </summary>
    public event EventHandler<IntPtr>? CursorChanged;

    /// <summary>
    /// Fired when focus state changes.
    /// </summary>
    public event EventHandler? FocusChanged;

    internal unsafe WebView2OffscreenView(
        WebView2OffscreenManager manager,
        StaThread staThread,
        ID3D11Device* device,
        CoreWebView2Environment environment,
        Compositor compositor,
        int width,
        int height)
    {
        _manager = manager;
        _staThread = staThread;
        _device = device;
        _environment = environment;
        _compositor = compositor;
        _width = width;
        _height = height;
    }

    /// <summary>
    /// Initializes the view. Must be called before use.
    /// </summary>
    public Task InitializeAsync()
    {
        if (_initialized)
            return Task.CompletedTask;

        return _staThread.RunAsync(async () =>
        {
            // Create a dedicated hidden window for this view
            // Each view needs its own window so EnumChildWindows finds the correct WebView2 child
            _hiddenWindow = new HiddenWindow();

            // Create root visual - this is what we'll capture from
            _rootVisual = _compositor.CreateContainerVisual();
            _rootVisual.Size = new System.Numerics.Vector2(_width, _height);
            _rootVisual.IsVisible = true;

            // Create a child visual for the WebView content
            _webViewVisual = _compositor.CreateContainerVisual();
            _webViewVisual.RelativeSizeAdjustment = new System.Numerics.Vector2(1.0f, 1.0f);
            _webViewVisual.IsVisible = true;

            // Add webview visual as child of root
            _rootVisual.Children.InsertAtTop(_webViewVisual);

            // Create composition controller with this view's hidden window
            _compositionController = await _environment.CreateCoreWebView2CompositionControllerAsync(_hiddenWindow.Handle);
            _controller = _compositionController as CoreWebView2Controller;

            if (_controller == null)
            {
                throw new InvalidOperationException("Failed to get CoreWebView2Controller from composition controller");
            }

            _webView = _controller.CoreWebView2;

            // Set bounds
            _controller.Bounds = new Rectangle(0, 0, _width, _height);
            _controller.IsVisible = true;

            // Enable transparent background for proper alpha compositing
            // This allows web pages with transparent/semi-transparent backgrounds to render correctly
            _controller.DefaultBackgroundColor = System.Drawing.Color.Transparent;

            // Set the webview visual as the target for the composition controller
            _compositionController.RootVisualTarget = _webViewVisual;

            // Configure WebView2 settings
            var settings = _webView.Settings;
            settings.AreDefaultContextMenusEnabled = false;
            settings.AreDevToolsEnabled = true;
            settings.IsStatusBarEnabled = false;
            settings.IsZoomControlEnabled = false;

            // Subscribe to events
            _webView.NavigationCompleted += OnNavigationCompleted;
            _compositionController.CursorChanged += OnCursorChanged;

            // Initialize cursor
            _cursor = _compositionController.Cursor;

            // Create capture helper and start capturing (unsafe operation)
            InitializeCaptureHelper();

            // Find WebView2's internal window for keyboard input
            // WebView2 creates child windows under the parent HWND
            FindWebViewChildWindow();

            _initialized = true;

            Ready?.Invoke(this, EventArgs.Empty);
        });
    }

    private unsafe void InitializeCaptureHelper()
    {
        _captureHelper = new GraphicsCaptureHelper(_device);
        _captureHelper.FrameCaptured += OnFrameCaptured;
        _captureHelper.StartCapture(_rootVisual!, _width, _height);
    }

    private unsafe void FindWebViewChildWindow()
    {
        if (_hiddenWindow == null)
            return;

        // WebView2 creates child windows with class names like "Chrome_WidgetWin_0"
        // We need to find this window to properly route keyboard input
        // Each view has its own hidden window, so EnumChildWindows only finds this view's children
        HWND foundHwnd = HWND.NULL;

        EnumChildWindows(_hiddenWindow.Handle, &EnumChildWindowsCallback, (LPARAM)(&foundHwnd));

        _webViewHwnd = foundHwnd;
    }

    [System.Runtime.InteropServices.UnmanagedCallersOnly]
    private static unsafe BOOL EnumChildWindowsCallback(HWND hwnd, LPARAM lParam)
    {
        char* className = stackalloc char[256];
        int length = GetClassNameW(hwnd, className, 256);

        if (length > 0)
        {
            var name = new string(className, 0, length);
            // Look for Chrome/WebView2 windows
            if (name.StartsWith("Chrome_") || name.Contains("WebView"))
            {
                *(HWND*)lParam = hwnd;
                return false; // Stop enumeration
            }
        }
        return true; // Continue enumeration
    }

    private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        NavigationCompleted?.Invoke(this, _webView?.Source ?? string.Empty);
    }

    private void OnCursorChanged(object? sender, object e)
    {
        if (_compositionController != null)
        {
            _cursor = _compositionController.Cursor;
            CursorChanged?.Invoke(this, _cursor);
            _manager.OnViewCursorChanged(this, _cursor);
        }
    }

    private void OnFrameCaptured(object? sender, IntPtr texturePtr)
    {
        FrameCaptured?.Invoke(this, texturePtr);
    }

    /// <summary>
    /// Sets focus to this view.
    /// </summary>
    public void SetFocus()
    {
        _manager.SetFocusedView(this);
    }

    /// <summary>
    /// Called by manager when this view gains focus.
    /// </summary>
    internal void OnFocusGained()
    {
        _isFocused = true;

        // Move WebView2's internal focus to this view
        // Note: We don't call Windows SetFocus() because we manually route keyboard messages
        // via PostMessageW. Calling SetFocus would cause double input.
        _staThread.Post(() =>
        {
            try
            {
                // Find the WebView2 child window for keyboard message routing
                if (_webViewHwnd == HWND.NULL)
                {
                    FindWebViewChildWindow();
                }

                // Tell WebView2 it has focus (for internal focus management like caret, selections)
                _controller?.MoveFocus(CoreWebView2MoveFocusReason.Programmatic);
            }
            catch
            {
                // Ignore errors during focus change
            }
        });

        FocusChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Called by manager when this view loses focus.
    /// </summary>
    internal void OnFocusLost()
    {
        _isFocused = false;
        FocusChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Navigates to the specified URL.
    /// </summary>
    public void Navigate(string url)
    {
        _staThread.Post(() => _webView?.Navigate(url));
    }

    /// <summary>
    /// Reloads the current page.
    /// </summary>
    public void Reload()
    {
        _staThread.Post(() => _webView?.Reload());
    }

    /// <summary>
    /// Opens the DevTools window.
    /// </summary>
    public void OpenDevTools()
    {
        _staThread.Post(() => _webView?.OpenDevToolsWindow());
    }

    /// <summary>
    /// Resizes the view to the specified dimensions.
    /// </summary>
    public void Resize(int width, int height)
    {
        if (width <= 0 || height <= 0)
            return;

        _width = width;
        _height = height;

        _staThread.Post(() =>
        {
            if (_controller != null)
            {
                _controller.Bounds = new Rectangle(0, 0, width, height);
            }

            if (_rootVisual != null)
            {
                _rootVisual.Size = new System.Numerics.Vector2(width, height);
            }

            _captureHelper?.Resize(width, height);
        });
    }

    #region Input Handling

    private static CoreWebView2MouseEventVirtualKeys ConvertVirtualKeys(VirtualKeys keys)
    {
        var result = CoreWebView2MouseEventVirtualKeys.None;

        if ((keys & VirtualKeys.Control) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.Control;

        if ((keys & VirtualKeys.Shift) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.Shift;

        if ((keys & VirtualKeys.LeftButton) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.LeftButton;

        if ((keys & VirtualKeys.RightButton) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.RightButton;

        if ((keys & VirtualKeys.MiddleButton) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.MiddleButton;

        if ((keys & VirtualKeys.XButton1) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.XButton1;

        if ((keys & VirtualKeys.XButton2) != 0)
            result |= CoreWebView2MouseEventVirtualKeys.XButton2;

        return result;
    }

    internal void HandleMouseMove(int x, int y, VirtualKeys modifiers)
    {
        if (_compositionController == null)
            return;

        var point = new Point(x, y);
        var keys = ConvertVirtualKeys(modifiers);
        _staThread.Post(() =>
            _compositionController?.SendMouseInput(CoreWebView2MouseEventKind.Move, keys, 0, point));
    }

    internal void HandleMouseButton(MouseButton button, bool down, int x, int y, VirtualKeys modifiers)
    {
        if (_compositionController == null)
            return;

        if (button == MouseButton.X1 || button == MouseButton.X2)
            return; // XButtons not supported by WebView2

        var point = new Point(x, y);
        var keys = ConvertVirtualKeys(modifiers);

        CoreWebView2MouseEventKind eventKind = button switch
        {
            MouseButton.Left => down ? CoreWebView2MouseEventKind.LeftButtonDown : CoreWebView2MouseEventKind.LeftButtonUp,
            MouseButton.Right => down ? CoreWebView2MouseEventKind.RightButtonDown : CoreWebView2MouseEventKind.RightButtonUp,
            MouseButton.Middle => down ? CoreWebView2MouseEventKind.MiddleButtonDown : CoreWebView2MouseEventKind.MiddleButtonUp,
            _ => CoreWebView2MouseEventKind.Move // XButtons not directly supported
        };

        _staThread.Post(() =>
            _compositionController?.SendMouseInput(eventKind, keys, 0, point));
    }

    internal void HandleMouseWheel(int delta, int x, int y, bool horizontal, VirtualKeys modifiers)
    {
        if (_compositionController == null)
            return;

        var point = new Point(x, y);
        var keys = ConvertVirtualKeys(modifiers);
        var eventKind = horizontal ? CoreWebView2MouseEventKind.HorizontalWheel : CoreWebView2MouseEventKind.Wheel;

        _staThread.Post(() =>
            _compositionController?.SendMouseInput(eventKind, keys, (uint)delta, point));
    }

    internal void HandleMouseLeave()
    {
        if (_compositionController == null)
            return;

        _staThread.Post(() =>
            _compositionController?.SendMouseInput(CoreWebView2MouseEventKind.Leave, CoreWebView2MouseEventVirtualKeys.None, 0, new Point(0, 0)));
    }

    internal void HandleKeyboardMessage(uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (_compositionController == null || _controller == null || _webView == null)
            return;

        _staThread.Post(() =>
        {
            // Try to find the WebView2 child window if not found yet
            if (_webViewHwnd == HWND.NULL)
            {
                FindWebViewChildWindow();
            }

            // If we found a WebView2 window, post message to it
            if (_webViewHwnd != HWND.NULL)
            {
                PostMessageW(_webViewHwnd, msg, (WPARAM)(nuint)wParam, (LPARAM)(nint)lParam);
            }
        });
    }

    #endregion

    /// <summary>
    /// Disposes the view asynchronously, ensuring proper cleanup on the STA thread.
    /// This is the preferred disposal method.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Remove from manager first
        _manager.RemoveView(this);

        // Run cleanup on STA thread and wait for completion
        await _staThread.RunAsync(() =>
        {
            DisposeCore();
            return Task.CompletedTask;
        });

        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes the view synchronously. Prefer DisposeAsync for proper cleanup.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Remove from manager
        _manager.RemoveView(this);

        // Try to run on STA thread if possible
        try
        {
            _staThread.Run(DisposeCore);
        }
        catch (ObjectDisposedException)
        {
            // STA thread may not be available, dispose directly
            DisposeCore();
        }

        GC.SuppressFinalize(this);
    }

    private void DisposeCore()
    {
        // Stop capture
        _captureHelper?.Dispose();

        // Unsubscribe from events
        if (_webView != null)
        {
            _webView.NavigationCompleted -= OnNavigationCompleted;
        }

        if (_compositionController != null)
        {
            _compositionController.CursorChanged -= OnCursorChanged;
        }

        // Close controller
        _controller?.Close();

        // Dispose visuals (but not compositor - it's shared)
        _webViewVisual?.Dispose();
        _rootVisual?.Dispose();

        // Dispose this view's hidden window
        _hiddenWindow?.Dispose();
    }
}
