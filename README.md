# Revit IFC Viewer

A GPU-accelerated 3D viewer for Revit projects that lets you open IFC models, sync the current Revit model, and inspect everything in one smooth external viewport.

The goal of this add-in is simple: make large project models easier to look at, navigate, compare, and understand without fighting Revit's heavy full-model perspective performance.

Instead of using Revit as the real-time 3D viewer, Revit IFC Viewer exports or loads model geometry into a Direct3D viewport, keeps the result linked back to IFC/Revit elements, and gives the user fast navigation, selection, properties, sectioning, opacity controls, and model comparison tools.

## What problem does this solve?

Revit is excellent for authoring BIM data, but full-scale 3D navigation can become uncomfortable on large projects. Perspective views may stutter, orbiting around big buildings can feel slow, and quickly comparing several discipline models can be frustrating.

Revit IFC Viewer is meant to be a lightweight coordination viewer next to Revit:

- Load one or many IFC models into a fast 3D renderer.
- Sync the active Revit 3D view into the same viewer.
- Navigate the whole building smoothly with orbit or first-person walk mode.
- Select elements and read their IFC/Revit properties.
- Hide, section, compare, and inspect model parts without changing the Revit model.
- Keep the viewer connected to Revit with manual sync, auto-sync, and follow-selection.

It is not meant to replace Revit. It is meant to make the visual checking and model-understanding part of project work faster and nicer.

## Who is this for?

- Electrical, low-voltage, HVAC, structural, architectural, and BIM users who need to review large models.
- Designers who want a smoother way to walk through a project while still keeping Revit open.
- Coordinators who need to compare exported IFC models with the live Revit model.
- Developers building Revit add-ins who want an example of a linked external 3D viewport with IFC and Revit geometry.

## Main workflow

1. Open Revit and launch **RK Tools → Tools → IFC Viewer**.
2. Load an IFC folder or individual IFC models.
3. Sync the active Revit 3D view, or pick a saved 3D view for repeatable exports.
4. Use the external viewer to orbit, walk, section, select elements, inspect properties, and compare models.
5. Enable auto-sync when you want Revit changes to update the viewer incrementally while you work.
6. Enable follow-selection when you want selecting an element in Revit to frame the matching object in the viewer.

## User-facing features

### Fast model viewing

- Opens IFC 2x3 and IFC 4 files through xBIM.
- Supports loading multiple IFC models into the same scene.
- Uses a GPU-accelerated Helix Toolkit SharpDX viewport.
- Uses a binary geometry cache so previously opened IFCs can load much faster.
- Watches loaded IFC files and shows a reload banner when a source file changes on disk.
- Suppresses overlapping duplicate geometry between loaded models to reduce z-fighting and visual noise.

### Revit-linked workflow

- Exports visible geometry from the active or selected Revit 3D view.
- Supports manual sync from Revit into the viewer.
- Supports auto-sync based on Revit document changes.
- Incremental sync re-exports only changed/deleted Revit elements where possible.
- Can follow Revit selection and frame the corresponding object in the viewer.
- Keeps Revit as the authoring tool and uses the viewer as a fast visual review surface.

### Navigation

- Orbit around the whole model or around a selected element.
- Pan with middle mouse or Shift + left drag.
- Zoom with the mouse wheel.
- First-person walk mode with WASD / arrow keys, Q/E vertical movement, Shift sprint, and right-drag mouse-look.
- Fit/reset camera tools for returning to the full model.
- Revit-style orientation cube with face, edge, and corner snapping.
- Compass ring that follows the camera heading.

### Selection and properties

- Click IFC or Revit elements in the viewport.
- Selected elements are highlighted with a separate overlay instead of changing the original model material.
- The properties panel shows element identity, type/name information, IFC GlobalId or Revit ElementId, and grouped parameters.
- Property groups are shown as tabs with an overflow menu for wide parameter sets.
- Clicking empty space clears selection.
- Spacebar hides/unhides the selected element.
- Unhide all restores hidden elements.

### Model comparison tools

- Load multiple IFCs and the live Revit model at the same time.
- Control IFC opacity and Revit opacity separately.
- Use wireframe overlay to see hard edges more clearly.
- Use section planes to cut through the scene without changing Revit views.
- Add up to four section planes by clicking model faces.
- Show/hide section plane rectangles while keeping the clip active.

### Viewer settings

- Floating non-modal settings window.
- Dark/light theme support.
- Live rendering settings without a separate Apply button.
- Window size and position persistence.
- Optional SSAO for a stronger visual depth effect when performance allows it.

## Developer overview

At a high level the project has two geometry pipelines that meet inside the same viewer:

```text
IFC files
  → xBIM parsing / tessellation
  → cached IFC model data
  → merged GPU meshes
  → Helix Toolkit SharpDX viewport

Revit 3D view
  → Revit CustomExporter
  → tessellated visible geometry
  → incremental patching when possible
  → merged GPU meshes
  → Helix Toolkit SharpDX viewport
```

The important design decision is that the viewer does **not** render one WPF/Helix node per Revit or IFC element. That approach becomes CPU-heavy very quickly on real models.

Instead, geometry is merged into larger renderable meshes, usually grouped by colour/transparency. Per-element identity is preserved separately through lightweight handles that remember each element's vertex/index ranges inside those merged meshes. That makes the scene much cheaper to draw while still allowing selection, hiding, zoom-to-element, property lookup, and Revit follow-selection.

## Repository structure

| Path | Purpose |
|---|---|
| `App.cs` | Revit external application, ribbon button registration, dependency resolution. |
| `Commands/` | Revit external command that opens or focuses the modeless viewer window. |
| `Core/` | Shared infrastructure such as session logging. |
| `Ifc/` | IFC loading, model records, caching, file watching, and xBIM-based geometry handling. |
| `IfcCore/` | IFC viewer service abstractions / shared IFC core layer. |
| `Revit/` | Revit geometry export, model records, incremental sync, external event bridge, selection following. |
| `UI/` | WPF viewer window, custom chrome, themes, controls, settings dialogs, overlay windows. |
| `Viewer/` | Helix viewport host, merged scene builder, camera logic, walk mode, section planes, wireframe, render settings. |
| `Deploy/` | Revit `.addin` manifests for supported versions. |

## How the rendering model works

### Merged mesh scene

Large BIM models can contain thousands of elements. Creating a separate 3D node for every wall, duct, cable tray, fitting, and family instance causes unnecessary CPU overhead from scene traversal, hit testing, state changes, frustum checks, and draw calls.

The viewer reduces that overhead by merging many element meshes into a small number of larger mesh buffers.

The key classes are:

- `MergedSceneBuilder` — collects per-element geometry and produces merged mesh buffers.
- `MergedMeshInfo` — stores the rendered mesh, original master index buffer, and element range metadata.
- `MergedPart` — describes one element's vertex/index range inside a merged mesh.
- `ElementHandle` — represents one logical IFC product or Revit element across one or more merged parts.

### Picking and element identity

Even though geometry is merged for performance, the viewer can still resolve a clicked triangle back to the correct source element.

The hit-test result gives a vertex/triangle location. The merged mesh metadata maps that vertex range back to an `ElementHandle`. The handle then points to IFC element info or Revit element info, which drives highlighting, properties, hide/unhide, zoom/orbit pivot, and follow-selection.

### Hide/unhide

Hidden elements are not removed from the master geometry buffers. Instead, the visible index buffer is rebuilt while skipping hidden element ranges. This keeps the original element mapping stable and allows fast restoration.

### Revit export and sync

Revit geometry is exported through the Revit API `CustomExporter` pipeline. A full export tessellates every visible element in the selected 3D view. Incremental sync tracks changed and deleted elements and re-tessellates only the dirty elements when possible, then patches the previous merged scene.

This means the viewer is view-based: it shows what the chosen Revit 3D view exports, not every hidden object in the entire Revit document.

### IFC loading and caching

IFC files are parsed through xBIM. The geometry is converted into the same viewer model format used by Revit exports. The viewer also keeps a binary geometry cache so repeated opens of the same unchanged IFC file can skip expensive processing.

Loaded IFC files are watched on disk. When a file changes, the viewer shows a reload notification so the user can refresh the model.

### Swap-chain rendering

The viewport uses Helix Toolkit SharpDX with swap-chain rendering. This avoids the slower WPF `D3DImage` presentation path and helps keep interaction smoother on large scenes.

Because the swap-chain render surface is a child HWND, normal WPF overlays cannot always draw above it. The orientation cube / compass is therefore hosted in a small owned overlay window that stays above the viewport.

## Build targets

| Configuration | Target framework | Revit version | Output folder |
|---|---:|---:|---|
| `Debug2024` / `Release2024` | `net48` | Revit 2024 | `bin/2024/` |
| `Debug2026` / `Release2026` | `net8.0-windows` | Revit 2026 | `bin/2026/` |

The project multi-targets Revit 2024 and Revit 2026 because the Revit API runtime changed from .NET Framework to modern .NET.

## Prerequisites

- Visual Studio 2022 or .NET SDK 8+.
- Autodesk Revit 2024 and/or Revit 2026 installed.
- Revit API DLLs available on the build machine.
- Windows x64.

The project probes the standard Revit install locations first. Custom Revit paths can be passed during build:

```powershell
dotnet build IfcViewer.csproj -c Debug2026 -p:Revit2026Dir="E:\Revit 2026"
dotnet build IfcViewer.csproj -c Debug2024 -p:Revit2024Dir="E:\Autodesk\Revit 2024"
```

Typical build command:

```powershell
dotnet build IfcViewer.csproj -c Debug2026
```

## Deployment

Copy the compiled add-in output and the matching `.addin` manifest from `Deploy/` into the Revit add-ins folder:

```text
%APPDATA%\Autodesk\Revit\Addins\2024\
%APPDATA%\Autodesk\Revit\Addins\2026\
```

The add-in uses Costura.Fody for embedded dependencies where possible. Some xBIM / Helix related assemblies are intentionally copied as loose DLLs next to `IfcViewer.dll` and loaded through the add-in assembly resolver.

## Key dependencies

| Package | Purpose |
|---|---|
| `Xbim.Essentials`, `Xbim.Geometry` | IFC parsing and geometry generation. |
| `HelixToolkit.Wpf.SharpDX` | Direct3D 11 viewport and GPU rendering. |
| `MaterialDesignThemes`, `MaterialDesignColors` | WPF control styling. |
| `Newtonsoft.Json` | Settings and persisted viewer state. |
| `ricaun.Revit.UI` | Revit ribbon helpers and app loading. |
| `Costura.Fody` | Dependency embedding for simpler deployment. |

## Current limitations / design notes

- This is a viewer and coordination tool, not an IFC editor.
- Revit sync is based on exported visible geometry from a 3D view.
- Section planes are viewer clip planes, not Revit section views.
- The viewer prioritizes smooth navigation and inspection over exact Revit visual parity.
- Some dependencies are embedded, while others are copied loose because of Revit runtime and package compatibility requirements.

## Roadmap ideas

Possible future improvements:

- Saved model comparison sessions per Revit project.
- Better colour/filter presets for disciplines or categories.
- Search by IFC GlobalId, Revit ElementId, name, type, category, or parameter value.
- Clash-style visual checks between loaded IFC and live Revit geometry.
- Issue/bookmark creation from selected viewer elements.
- More advanced section box / clipping workflows.
- Packaging installer for easier team deployment.

## Project goal

The long-term goal is to make model review feel lightweight: open the project, load the relevant IFCs, sync Revit, and move around the whole building smoothly while still being able to select real elements and understand what they are.

In other words, Revit remains the source of truth, but the viewer becomes the fast visual workspace for understanding the model.