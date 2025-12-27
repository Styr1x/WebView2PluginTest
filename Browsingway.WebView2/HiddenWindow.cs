using System;
using System.Runtime.InteropServices;
using System.Threading;
using TerraFX.Interop.Windows;
using static TerraFX.Interop.Windows.Windows;

namespace Browsingway.WebView2;

/// <summary>
/// Creates a message-only window for WebView2 composition controller.
/// This window is not visible and is used only for hosting the composition visual tree.
/// </summary>
internal sealed unsafe class HiddenWindow : IDisposable
{
    private static readonly Lock ClassLock = new();
    private static bool _classRegistered;
    private static ushort _classAtom;
    private static delegate* unmanaged<HWND, uint, WPARAM, LPARAM, LRESULT> _wndProcPtr;

    // HWND_MESSAGE constant - message-only window parent
    private static readonly HWND HWND_MESSAGE = new HWND((void*)-3);

    private const string ClassName = "Browsingway_HiddenWindow";
    private const string WindowName = "Browsingway.WebView2";

    private bool _disposed;

    public HWND Handle { get; private set; }

    public HiddenWindow()
    {
        EnsureClassRegistered();
        CreateHiddenWindow();
    }

    private static void EnsureClassRegistered()
    {
        lock (ClassLock)
        {
            if (_classRegistered)
                return;

            var hInstance = GetModuleHandleW(null);

            // Use function pointer for the window procedure
            _wndProcPtr = &DefWndProcHandler;

            fixed (char* classNamePtr = ClassName)
            {
                var wc = new WNDCLASSW
                {
                    style = 0,
                    lpfnWndProc = _wndProcPtr,
                    cbClsExtra = 0,
                    cbWndExtra = 0,
                    hInstance = hInstance,
                    hIcon = HICON.NULL,
                    hCursor = HCURSOR.NULL,
                    hbrBackground = HBRUSH.NULL,
                    lpszMenuName = null,
                    lpszClassName = classNamePtr
                };

                _classAtom = RegisterClassW(&wc);
                if (_classAtom == 0)
                {
                    int error = Marshal.GetLastWin32Error();
                    // Class may already exist from previous run, try to get it
                    WNDCLASSW existingClass;
                    if (GetClassInfoW(hInstance, classNamePtr, &existingClass) != 0)
                    {
                        _classRegistered = true;
                        return;
                    }
                    throw new InvalidOperationException($"Failed to register window class. Error: {error}");
                }
            }

            _classRegistered = true;
        }
    }

    [UnmanagedCallersOnly]
    private static LRESULT DefWndProcHandler(HWND hWnd, uint msg, WPARAM wParam, LPARAM lParam)
    {
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private void CreateHiddenWindow()
    {
        var hInstance = GetModuleHandleW(null);

        fixed (char* windowNamePtr = WindowName)
        {
            // Create a message-only window (parent = HWND_MESSAGE)
            Handle = CreateWindowExW(
                0,
                (char*)_classAtom,
                windowNamePtr,
                0,
                0, 0, 0, 0,
                HWND_MESSAGE,
                HMENU.NULL,
                hInstance,
                null
            );
        }

        if (Handle == HWND.NULL)
        {
            int error = Marshal.GetLastWin32Error();
            throw new InvalidOperationException($"Failed to create hidden window. Error: {error}");
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        if (Handle != HWND.NULL)
        {
            DestroyWindow(Handle);
            Handle = HWND.NULL;
        }
    }
}
