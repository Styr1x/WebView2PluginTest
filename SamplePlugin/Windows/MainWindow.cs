using System;
using System.Numerics;
using Dalamud.Bindings.ImGui;
using Dalamud.Interface.Utility;
using Dalamud.Interface.Utility.Raii;
using Dalamud.Interface.Windowing;
using Dalamud.IoC;
using Dalamud.Plugin.Services;
using Lumina.Excel.Sheets;
using TerraFX.Interop.DirectX;

namespace SamplePlugin.Windows;

public class MainWindow : Window, IDisposable
{
    [PluginService]
    internal static IPluginLog Log { get; private set; } = null!;

    private readonly string goatImagePath;
    private readonly Plugin plugin;
    private IntPtr testTexture;
    private ImTextureID _textureId;
    private IPluginLog _log;

    // We give this window a hidden ID using ##.
    // The user will see "My Amazing Window" as window title,
    // but for ImGui the ID is "My Amazing Window##With a hidden ID"
    public MainWindow(Plugin plugin, string goatImagePath, IPluginLog log)
        : base("My Amazing Window##With a hidden ID", ImGuiWindowFlags.NoScrollbar | ImGuiWindowFlags.NoScrollWithMouse)
    {
        SizeConstraints = new WindowSizeConstraints
        {
            MinimumSize = new Vector2(375, 330),
            MaximumSize = new Vector2(float.MaxValue, float.MaxValue)
        };

        this.goatImagePath = goatImagePath;
        this.plugin = plugin;
        this._log = log;
    }

    public unsafe void SetTestTexture(ID3D11Texture2D* tex, ID3D11Device* device)
    {
        if (testTexture != IntPtr.Zero) return;
        
        lock (this)
        {
            tex->AddRef();
            ID3D11ShaderResourceView* srv;
            int hr = device->CreateShaderResourceView((ID3D11Resource*)tex, null, &srv);
            if (hr >= 0)
            {
                testTexture = (IntPtr)srv;
                _textureId = new ImTextureID(testTexture);
                _log.Information("Created test texture.");
            }
            else
            {
                _log.Error($"Failed to create test texture: {hr}");
                return;
            }
        }
    }

    public void Dispose() { }

    public override void Draw()
    {
        ImGui.TextUnformatted($"The random config bool is {plugin.Configuration.SomePropertyToBeSavedAndWithADefault}");

        if (ImGui.Button("Show Settings"))
        {
            plugin.ToggleConfigUi();
        }

        ImGui.Spacing();

        // Normally a BeginChild() would have to be followed by an unconditional EndChild(),
        // ImRaii takes care of this after the scope ends.
        // This works for all ImGui functions that require specific handling, examples are BeginTable() or Indent().
        using (var child = ImRaii.Child("SomeChildWithAScrollbar", Vector2.Zero, true))
        {
            // Check if this child is drawing
            if (child.Success)
            {
                ImGui.TextUnformatted("Have a goat:");
                var goatImage = Plugin.TextureProvider.GetFromFile(goatImagePath).GetWrapOrDefault();
                if (goatImage != null)
                {
                    using (ImRaii.PushIndent(55f))
                    {
                        ImGui.Image(goatImage.Handle, goatImage.Size);
                    }
                }
                else
                {
                    ImGui.TextUnformatted("Image not found.");
                }

                ImGuiHelpers.ScaledDummy(20.0f);

                lock (this)
                {
                    Vector2 _size = new Vector2(1400, 1400);
                    ImGui.Image(_textureId, _size);
                }

                // Example for other services that Dalamud provides.
                // ClientState provides a wrapper filled with information about the local player object and client.

                var localPlayer = Plugin.ClientState.LocalPlayer;
                if (localPlayer == null)
                {
                    ImGui.TextUnformatted("Our local player is currently not loaded.");
                    return;
                }

                if (!localPlayer.ClassJob.IsValid)
                {
                    ImGui.TextUnformatted("Our current job is currently not valid.");
                    return;
                }

                // If you want to see the Macro representation of this SeString use `ToMacroString()`
                ImGui.TextUnformatted($"Our current job is ({localPlayer.ClassJob.RowId}) \"{localPlayer.ClassJob.Value.Abbreviation}\"");

                // Example for quarrying Lumina directly, getting the name of our current area.
                var territoryId = Plugin.ClientState.TerritoryType;
                if (Plugin.DataManager.GetExcelSheet<TerritoryType>().TryGetRow(territoryId, out var territoryRow))
                {
                    ImGui.TextUnformatted($"We are currently in ({territoryId}) \"{territoryRow.PlaceName.Value.Name}\"");
                }
                else
                {
                    ImGui.TextUnformatted("Invalid territory.");
                }
            }
        }
    }
}
