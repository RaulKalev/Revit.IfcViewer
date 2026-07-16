# Loader Harness

Console harness that exercises the built `IfcViewer.dll` **outside Revit** —
it simulates Revit's environment (compile-time references only, no dependency
files next to the exe) and verifies the two riskiest subsystems end to end:

1. **Single-DLL packaging** — loads the plugin, lets `DependencyLoader`
   extract the embedded libs to `%LOCALAPPDATA%`, and resolves every managed
   assembly plus the native xBIM/OpenCascade engine through its hooks.
2. **IFC pipeline** — runs the plugin's real `IfcLoader.LoadAsync` on a WPF
   dispatcher thread: xBIM v6 tessellation, opening cuts, Helix scene build,
   binary-cache write, then a second load through the cache-hit path, and
   checks both loads produce identical triangle counts.

## Usage

```
dotnet build IfcViewer.csproj -c Release2026
dotnet build Tests/LoaderHarness -c Release

# full plugin loader test (cold + cached load)
Tests\LoaderHarness\bin\Release\net8.0-windows\LoaderHarness.exe ^
    bin\Release2026\net8.0-windows\IfcViewer.dll ^
    Tests\LoaderHarness\wall-with-opening.ifc loader

# raw geometry-engine smoke test
Tests\LoaderHarness\bin\Release\net8.0-windows\LoaderHarness.exe ^
    bin\Release2026\net8.0-windows\IfcViewer.dll ^
    Tests\LoaderHarness\wall.ifc engine
```

Use the `net48` harness build against `Release2024` output for the Revit 2024
target. Exit code 0 = pass. Any real-world `.ifc` file can be substituted for
the bundled minimal ones (`wall.ifc`: one extruded wall, 12 triangles;
`wall-with-opening.ifc`: same wall voided by an `IfcOpeningElement`, 28
triangles).

The harness is intentionally **not** part of `IfcViewer.sln` so plugin builds
stay fast; build it directly by path as shown above.
