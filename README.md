# IFC Viewer

A Revit add-in that provides a high-performance, interactive 3D viewer for IFC files and live Revit geometry — built on [xBIM](https://xbimteam.github.io/) and [Helix Toolkit SharpDX](https://github.com/helix-toolkit/helix-toolkit).

## Features

### IFC Loading
- Loads IFC 2x3 and IFC 4 files via xBIM geometry engine
- Multi-layer wall support — compound wall structures are merged into single buckets, avoiding doubled geometry
- Door and window openings cut through host walls using AABB plane-splitting for clean straight edges
- Binary geometry cache for instant subsequent opens (invalidated automatically on file change)
- Auto-reload banner when the source IFC file changes on disk

### Revit Integration
- Export the active 3D view (or a chosen saved 3D view) directly into the viewer
- Incremental sync — only changed/deleted elements are re-exported on `DocumentChanged`
- Auto-sync toggle for continuous live updates as you model
- Per-project saved 3D view preference

### Navigation
- **Orbit** — left-drag to rotate around the scene or around a selected element (Revit-style pivot)
- **Pan** — middle-drag or Shift+left-drag
- **Zoom** — scroll wheel
- **Walk mode** — WASD / arrow keys, Q/E vertical, Shift to sprint, right-drag mouse-look
- **Fit to view** — fits camera to the full scene on first load; retains position on subsequent loads

### Orientation Cube (ViewCube)
- Fully custom Revit-style cube in the top-right corner with FRONT/BACK/LEFT/RIGHT/TOP/BOTTOM labels
- Click a face, edge, or corner to snap to that isometric or orthographic view
- Compass ring below the cube rotates with the camera heading
- Rendered in a small owned overlay window so it stays visible above the swap-chain viewport (see Rendering & Performance)

### Element Selection & Properties
- Click any IFC or Revit element to select it; teal emissive highlight overlay on selection
- Selection pivots the orbit camera around the element's bounding-box centre
- Properties panel shows element identity (Name, Type, GlobalId / ElementId) and all parameter groups
- Tabbed parameter groups with chevron overflow menu for wide screens
- Clicking a selected element deselects it; clicking empty space clears selection
- Spacebar hides the selected element (and un-hides on repeat); works in both orbit and walk mode

### Rendering & Performance
- Geometry is merged by colour into a handful of large meshes per model instead of one mesh per element — a model with thousands of IFC products renders as a few dozen draw calls, which is what keeps interaction smooth. Per-element identity (selection, hide/unhide, zoom-to, Revit follow-selection) is preserved via vertex-range handles with a hit-test octree per merged mesh, so picking stays fast even on multi-million-triangle scenes
- DXGI swap-chain rendering (`Viewport3DX.EnableSwapChainRendering`) instead of the WPF `D3DImage` path — presents vsynced directly to the compositor, eliminating the frame-rate ceiling and screen tearing that D3DImage's unsynchronized shared-surface copy caused
- SSAO is off by default (it costs several full-resolution GPU passes per frame) — toggle it back on under Settings → Rendering if you want the ambient-occlusion look and can spare the frame budget
- Hard-edge wireframe overlay (dihedral-angle filtered, ~30° threshold) — toggled per session
- FXAA anti-aliasing
- Section plane — arbitrary GPU clip plane with a magenta transparent visual quad, set by right-clicking a face

### Settings
- Floating settings dialog (non-modal, dark theme)
- All sliders/toggles apply live — no apply/save button needed
- Window position and size persisted across sessions

## Architecture

| Layer | Project path | Responsibility |
|---|---|---|
| Core | `Core/` | Session logging |
| IFC | `Ifc/` | xBIM loading, meshing, caching |
| Revit | `Revit/` | CustomExporter, incremental sync, ExternalEvent bridge |
| Viewer | `Viewer/` | Helix scene, camera, settings, walk controller, section plane, merged-mesh scene building |
| UI | `UI/` | WPF windows, custom chrome, styles, swap-chain input bridge, viewport overlay window |

## Build Targets

| Configuration | Target Framework | Revit Version | Output |
|---|---|---|---|
| Debug2024 / Release2024 | net48 | Revit 2024 | `DevDlls/IfcViewer/2024/` |
| Debug2026 / Release2026 | net8.0-windows | Revit 2026 | `DevDlls/IfcViewer/2026/` |

Both targets compile cleanly. Revit API DLLs are referenced as local files (not NuGet) and must be present on the build machine.

### Prerequisites
- Visual Studio 2022 or `dotnet` SDK 8+
- Revit 2024 API DLLs at `E:\Autodesk\Revit 2024\`
- Revit 2026 API DLLs at `E:\Revit 2026\`

```
dotnet build IfcViewer.csproj -c Debug2026
```

## Deployment

Copy the build output and the matching `.addin` manifest from `Deploy/` to the Revit addins folder:

- Revit 2024: `%APPDATA%\Autodesk\Revit\Addins\2024\`
- Revit 2026: `%APPDATA%\Autodesk\Revit\Addins\2026\`

The DLL is self-contained — xBIM and all other dependencies are embedded via Costura.Fody.

## Key Dependencies

| Package | Purpose |
|---|---|
| Xbim.Essentials, Xbim.Geometry | IFC parsing and 3D geometry engine |
| HelixToolkit.Wpf.SharpDX | Direct3D 11 GPU viewport |
| Newtonsoft.Json | Settings serialisation |
| MaterialDesignThemes | WPF control styles |
| Costura.Fody | Embed all dependencies into a single DLL |
