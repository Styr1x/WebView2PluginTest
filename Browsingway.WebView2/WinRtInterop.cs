using System;
using System.Runtime.InteropServices;
using TerraFX.Interop.DirectX;
using TerraFX.Interop.Windows;
using Windows.Graphics.DirectX.Direct3D11;
using WinRT;

namespace Browsingway.WebView2;

/// <summary>
/// Helper class for WinRT and Direct3D11 interop operations.
/// Provides utilities for converting between WinRT and native D3D11 types.
/// </summary>
internal static unsafe class WinRtInterop
{
    // COM interface for accessing the underlying DXGI interface from WinRT IDirect3DSurface
    [ComImport]
    [Guid("A9B3D012-3DF2-4EE3-B8D1-8695F457D3C1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDirect3DDxgiInterfaceAccess
    {
        [PreserveSig]
        HRESULT GetInterface([In] ref Guid iid, out IntPtr ppv);
    }

    private static readonly Guid ID3D11Texture2DGuid = new("6f15aaf2-d208-4e89-9ab4-489535d34f9c");
    private static readonly Guid IDirect3DDxgiInterfaceAccessGuid = new("A9B3D012-3DF2-4EE3-B8D1-8695F457D3C1");



    [DllImport("d3d11.dll", EntryPoint = "CreateDirect3D11DeviceFromDXGIDevice", SetLastError = true)]
    private static extern int CreateDirect3D11DeviceFromDXGIDevice(
        IntPtr dxgiDevice,
        out IntPtr graphicsDevice);

    /// <summary>
    /// Extracts the native ID3D11Texture2D pointer from a WinRT IDirect3DSurface.
    /// </summary>
    /// <param name="surface">The WinRT surface to extract from.</param>
    /// <returns>The native texture pointer, or IntPtr.Zero if extraction fails.</returns>
    /// <remarks>
    /// The caller is responsible for releasing the returned texture pointer.
    /// </remarks>
    public static IntPtr GetTextureFromSurface(IDirect3DSurface surface)
    {
        // The IDirect3DSurface in CsWinRT is an IWinRTObject that wraps the native object
        // We need to get the native pointer and query for IDirect3DDxgiInterfaceAccess from there
        if (surface is IWinRTObject winrtObj)
        {
            var nativePtr = winrtObj.NativeObject.ThisPtr;

            // Query for IDirect3DDxgiInterfaceAccess from the native object
            var accessGuid = IDirect3DDxgiInterfaceAccessGuid;
            HRESULT hr = Marshal.QueryInterface(nativePtr, in accessGuid, out IntPtr accessPtr);

            if (hr.SUCCEEDED && accessPtr != IntPtr.Zero)
            {
                try
                {
                    // Use the COM interface instead of manual vtable indexing
                    var access = (IDirect3DDxgiInterfaceAccess)Marshal.GetObjectForIUnknown(accessPtr);

                    var textureGuid = ID3D11Texture2DGuid;
                    hr = access.GetInterface(ref textureGuid, out IntPtr texturePtr);

                    if (hr.SUCCEEDED)
                    {
                        return texturePtr;
                    }
                }
                finally
                {
                    Marshal.Release(accessPtr);
                }
            }
        }

        return IntPtr.Zero;
    }

    /// <summary>
    /// Creates a WinRT IDirect3DDevice from a native D3D11 device.
    /// </summary>
    /// <param name="device">The native D3D11 device pointer.</param>
    /// <returns>The WinRT IDirect3DDevice wrapper.</returns>
    public static IDirect3DDevice CreateWinRTDevice(ID3D11Device* device)
    {
        // Get DXGI device from D3D11 device
        IDXGIDevice* dxgiDevice;
        Guid dxgiGuid = typeof(IDXGIDevice).GUID;
        HRESULT hr = ((IUnknown*)device)->QueryInterface(&dxgiGuid, (void**)&dxgiDevice);
        Marshal.ThrowExceptionForHR(hr);

        try
        {
            // Create WinRT Direct3D device from DXGI device
            hr = CreateDirect3D11DeviceFromDXGIDevice((IntPtr)dxgiDevice, out IntPtr inspectable);

            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            // Wrap the inspectable in WinRT IDirect3DDevice
            return MarshalInterface<IDirect3DDevice>.FromAbi(inspectable);
        }
        finally
        {
            dxgiDevice->Release();
        }
    }
}
