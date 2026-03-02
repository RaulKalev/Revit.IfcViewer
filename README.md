# IFC Viewer

A Revit add-in that provides a high-performance, interactive 3D viewer for IFC files and live Revit geometry — built on [xBIM](https://xbimteam.github.io/) and [Helix Toolkit SharpDX](https://github.com/helix-toolkit/helix-toolkit).

## Features

### IFC Loading
- Loads IFC 2x3 and IFC 4 files via xBIM geometry engine
- Per-element mesh extraction with correct material colours
- Multi-layer wall support — compound wall structures are merged into single buckets, avoiding doubled geometry
- Door and window openings cut through host walls using AABB plane-splitting for clean straight edges
- Progressive streaming into the viewport during load
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
- Revit-style ViewCube in the top-right corner with FRONT/BACK/LEFT/RIGHT/TOP/BOTTOM labels
- Click a face, edge, or corner to snap to that isometric or orthographic view
- Compass ring below the cube rotates with the camera heading

### Element Selection & Properties
- Click any IFC or Revit element to select it; teal emissive highlight on selection
- Selection pivots the orbit camera around the element's bounding-box centre
- Properties panel shows element identity (Name, Type, GlobalId / ElementId) and all parameter groups
- Tabbed parameter groups with chevron overflow menu for wide screens
- Clicking a selected element deselects it; clicking empty space clears selection

### Rendering
- Hard-edge wireframe overlay (dihedral-angle filtered, ~30° threshold) — toggled per session
- Always-on element outline (60° threshold, thinner lines) for silhouette clarity
- SSAO with building-scale sampling radius
- FXAA anti-aliasing
- Section plane — axis-aligned GPU clip plane with blue transparent visual quad

### Settings
- Floating settings dialog (non-modal, dark theme)
- All sliders apply live — no apply/save button needed
- Window position and size persisted across sessions

## Architecture

| Layer | Project path | Responsibility |
|---|---|---|
| Core | `Core/` | Session logging |
| IFC | `Ifc/` | xBIM loading, meshing, caching |
| Revit | `Revit/` | CustomExporter, incremental sync, ExternalEvent bridge |
| Viewer | `Viewer/` | Helix scene, camera, settings, walk controller, section plane |
| UI | `UI/` | WPF windows, custom chrome, styles |

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
