using System;

namespace Browsingway.WebView2;

/// <summary>
/// Represents modifier keys and mouse button states for input events.
/// </summary>
[Flags]
public enum VirtualKeys
{
    None = 0,
    LeftButton = 1 << 0,
    RightButton = 1 << 1,
    MiddleButton = 1 << 2,
    XButton1 = 1 << 3,
    XButton2 = 1 << 4,
    Shift = 1 << 5,
    Control = 1 << 6,
    Alt = 1 << 7
}
