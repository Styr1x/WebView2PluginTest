using System;
using System.Runtime.InteropServices;
using TerraFX.Interop.DirectX;
using TerraFX.Interop.Windows;
using Windows.Graphics;
using Windows.Graphics.Capture;
using Windows.Graphics.DirectX;
using Windows.Graphics.DirectX.Direct3D11;
using Windows.UI.Composition;

namespace Browsingway.WebView2;

/// <summary>
/// Captures frames from a Visual using Windows.Graphics.Capture API
/// and provides them as Direct3D11 textures.
/// </summary>
internal sealed unsafe class GraphicsCaptureHelper : IDisposable
{
    private readonly ID3D11Device* _device;
    private readonly ID3D11DeviceContext* _context;
    private IDirect3DDevice? _winrtDevice;
    private Direct3D11CaptureFramePool? _framePool;
    private GraphicsCaptureSession? _session;
    private GraphicsCaptureItem? _item;
    private SizeInt32 _lastSize;
    private ID3D11Texture2D* _capturedTexture;
    private ID3D11Texture2D* _stagingTexture; // CPU-readable copy for alpha sampling
    private byte[]? _pixelData; // Cached pixel data for alpha lookups (double-buffered)
    private byte[]? _pixelDataBack; // Back buffer for async update
    private int _pixelDataWidth;
    private int _pixelDataHeight;
    private int _framesSinceLastAlphaUpdate;
    private const int AlphaUpdateInterval = 5; // Update alpha data every N frames to reduce GPU stalls
    private bool _disposed;
    private readonly object _lock = new();
    private readonly object _pixelDataLock = new(); // Separate lock for non-blocking pixel data access

    public event EventHandler<IntPtr>? FrameCaptured;

    public ID3D11Texture2D* CapturedTexture
    {
        get
        {
            lock (_lock)
            {
                return _capturedTexture;
            }
        }
    }

    public GraphicsCaptureHelper(ID3D11Device* device)
    {
        _device = device;
        ID3D11DeviceContext* ctx;
        _device->GetImmediateContext(&ctx);
        _context = ctx;
        _winrtDevice = WinRtInterop.CreateWinRTDevice(_device);
    }

    public void StartCapture(Visual visual, int width, int height)
    {
        if (_winrtDevice == null)
            throw new InvalidOperationException("WinRT device not initialized");

        StopCapture();

        // Create capture item from visual
        _item = GraphicsCaptureItem.CreateFromVisual(visual);
        _lastSize = new SizeInt32 { Width = width, Height = height };

        // Create frame pool
        _framePool = Direct3D11CaptureFramePool.Create(
            _winrtDevice,
            DirectXPixelFormat.B8G8R8A8UIntNormalized,
            2,
            _lastSize);

        _framePool.FrameArrived += OnFrameArrived;

        // Create capture session
        _session = _framePool.CreateCaptureSession(_item);

        // Disable cursor capture and border (Windows 11 feature)
        try
        {
            // These properties may not exist on older Windows versions
            var sessionType = _session.GetType();
            var cursorProperty = sessionType.GetProperty("IsCursorCaptureEnabled");
            cursorProperty?.SetValue(_session, false);

            var borderProperty = sessionType.GetProperty("IsBorderRequired");
            borderProperty?.SetValue(_session, false);
        }
        catch
        {
            // Ignore if properties don't exist
        }

        // Start capturing
        _session.StartCapture();
    }

    public void StopCapture()
    {
        _session?.Dispose();
        _session = null;

        if (_framePool != null)
        {
            _framePool.FrameArrived -= OnFrameArrived;
            _framePool.Dispose();
            _framePool = null;
        }

        _item = null;
    }

    public void Resize(int width, int height)
    {
        if (_framePool == null || _winrtDevice == null)
            return;

        _lastSize = new SizeInt32 { Width = width, Height = height };
        _framePool.Recreate(_winrtDevice, DirectXPixelFormat.B8G8R8A8UIntNormalized, 2, _lastSize);
    }

    private void OnFrameArrived(Direct3D11CaptureFramePool sender, object args)
    {
        using var frame = sender.TryGetNextFrame();
        if (frame == null)
            return;

        if (frame.ContentSize.Width == 0 || frame.ContentSize.Height == 0)
            return;

        // Check if size changed
        if (frame.ContentSize.Width != _lastSize.Width || frame.ContentSize.Height != _lastSize.Height)
        {
            _lastSize = frame.ContentSize;
            _framePool?.Recreate(_winrtDevice!, DirectXPixelFormat.B8G8R8A8UIntNormalized, 2, _lastSize);
            return;
        }

        // Get the texture from the frame
        var surface = frame.Surface;

        try
        {
            // Get the D3D11 texture from the capture surface using the interop interface
            var texturePtr = WinRtInterop.GetTextureFromSurface(surface);

            if (texturePtr == IntPtr.Zero)
                return;

            // Get source texture as TerraFX type
            var sourceTexture = (ID3D11Texture2D*)texturePtr;

            lock (_lock)
            {
                // Get source texture description
                D3D11_TEXTURE2D_DESC sourceDesc;
                sourceTexture->GetDesc(&sourceDesc);

                // Create or recreate our texture if needed
                if (_capturedTexture == null)
                {
                    CreateCapturedTexture(sourceDesc.Width, sourceDesc.Height);
                }
                else
                {
                    D3D11_TEXTURE2D_DESC capturedDesc;
                    _capturedTexture->GetDesc(&capturedDesc);

                    if (capturedDesc.Width != sourceDesc.Width || capturedDesc.Height != sourceDesc.Height)
                    {
                        _capturedTexture->Release();
                        _capturedTexture = null;
                        CreateCapturedTexture(sourceDesc.Width, sourceDesc.Height);
                    }
                }

                // Copy the frame to our texture
                _context->CopyResource((ID3D11Resource*)_capturedTexture, (ID3D11Resource*)sourceTexture);

                // Periodically update alpha data for click-through detection
                // Don't do this every frame to avoid GPU stalls
                _framesSinceLastAlphaUpdate++;
                if (_framesSinceLastAlphaUpdate >= AlphaUpdateInterval && _stagingTexture != null)
                {
                    _framesSinceLastAlphaUpdate = 0;
                    UpdatePixelDataCache();
                }
            }

            // Release our reference to the source texture
            sourceTexture->Release();

            FrameCaptured?.Invoke(this, texturePtr);
        }
        catch
        {
            // Ignore capture errors - frame will be dropped
        }
    }

    private void CreateCapturedTexture(uint width, uint height)
    {
        var desc = new D3D11_TEXTURE2D_DESC
        {
            Width = width,
            Height = height,
            MipLevels = 1,
            ArraySize = 1,
            Format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
            SampleDesc = new DXGI_SAMPLE_DESC { Count = 1, Quality = 0 },
            Usage = D3D11_USAGE.D3D11_USAGE_DEFAULT,
            BindFlags = (uint)D3D11_BIND_FLAG.D3D11_BIND_SHADER_RESOURCE,
            CPUAccessFlags = 0,
            MiscFlags = 0
        };

        ID3D11Texture2D* texture;
        int hr = _device->CreateTexture2D(&desc, null, &texture);
        Marshal.ThrowExceptionForHR(hr);
        _capturedTexture = texture;

        // Also create/recreate staging texture for CPU read access
        CreateStagingTexture(width, height);
    }

    private void CreateStagingTexture(uint width, uint height)
    {
        // Release old staging texture if exists
        if (_stagingTexture != null)
        {
            _stagingTexture->Release();
            _stagingTexture = null;
        }

        var stagingDesc = new D3D11_TEXTURE2D_DESC
        {
            Width = width,
            Height = height,
            MipLevels = 1,
            ArraySize = 1,
            Format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
            SampleDesc = new DXGI_SAMPLE_DESC { Count = 1, Quality = 0 },
            Usage = D3D11_USAGE.D3D11_USAGE_STAGING,
            BindFlags = 0,
            CPUAccessFlags = (uint)D3D11_CPU_ACCESS_FLAG.D3D11_CPU_ACCESS_READ,
            MiscFlags = 0
        };

        ID3D11Texture2D* staging;
        int hr = _device->CreateTexture2D(&stagingDesc, null, &staging);
        if (hr >= 0)
        {
            _stagingTexture = staging;
        }

        // Reset pixel data cache
        lock (_pixelDataLock)
        {
            _pixelData = null;
            _pixelDataBack = null;
            _pixelDataWidth = 0;
            _pixelDataHeight = 0;
        }
    }

    /// <summary>
    /// Gets the alpha value at the specified pixel position.
    /// Returns 255 (opaque) if position is out of bounds or data unavailable.
    /// This method is non-blocking and uses cached pixel data.
    /// </summary>
    public byte GetAlphaAtPosition(int x, int y)
    {
        // Use separate lock to avoid blocking frame capture
        lock (_pixelDataLock)
        {
            if (_pixelData == null || x < 0 || y < 0 || x >= _pixelDataWidth || y >= _pixelDataHeight)
                return 255;

            // B8G8R8A8 format: each pixel is 4 bytes (B, G, R, A)
            int pixelIndex = (y * _pixelDataWidth + x) * 4;
            if (pixelIndex + 3 >= _pixelData.Length)
                return 255;

            // Alpha is the 4th byte (index 3) in B8G8R8A8
            return _pixelData[pixelIndex + 3];
        }
    }

    private void UpdatePixelDataCache()
    {
        if (_capturedTexture == null || _stagingTexture == null)
            return;

        // Copy captured texture to staging texture
        _context->CopyResource((ID3D11Resource*)_stagingTexture, (ID3D11Resource*)_capturedTexture);

        // Map the staging texture for CPU read (use DO_NOT_WAIT to avoid blocking)
        D3D11_MAPPED_SUBRESOURCE mapped;
        HRESULT hr = _context->Map((ID3D11Resource*)_stagingTexture, 0, D3D11_MAP.D3D11_MAP_READ,
            0, &mapped);
        if (hr < 0)
        {
            // GPU not ready, skip this update
            return;
        }

        try
        {
            D3D11_TEXTURE2D_DESC desc;
            _stagingTexture->GetDesc(&desc);

            int width = (int)desc.Width;
            int height = (int)desc.Height;
            int dataSize = width * height * 4;

            // Use back buffer for writing to avoid blocking reads
            if (_pixelDataBack == null || _pixelDataBack.Length != dataSize)
            {
                _pixelDataBack = new byte[dataSize];
            }

            // Copy row by row (accounting for pitch)
            byte* srcPtr = (byte*)mapped.pData;
            for (int row = 0; row < height; row++)
            {
                int dstOffset = row * width * 4;
                Marshal.Copy((IntPtr)(srcPtr + row * mapped.RowPitch), _pixelDataBack, dstOffset, width * 4);
            }

            // Swap buffers under lock
            lock (_pixelDataLock)
            {
                (_pixelData, _pixelDataBack) = (_pixelDataBack, _pixelData);
                _pixelDataWidth = width;
                _pixelDataHeight = height;
            }
        }
        finally
        {
            _context->Unmap((ID3D11Resource*)_stagingTexture, 0);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        StopCapture();

        lock (_lock)
        {
            if (_capturedTexture != null)
            {
                _capturedTexture->Release();
                _capturedTexture = null;
            }

            if (_stagingTexture != null)
            {
                _stagingTexture->Release();
                _stagingTexture = null;
            }
        }

        lock (_pixelDataLock)
        {
            _pixelData = null;
            _pixelDataBack = null;
        }

        if (_context != null)
        {
            _context->Release();
        }

        _winrtDevice?.Dispose();
    }
}