using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine($"Runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}, x64={Environment.Is64BitProcess}");
        string dllPath = Path.GetFullPath(args[0]);
        string ifcPath = Path.GetFullPath(args[1]);
        string mode = args.Length > 2 ? args[2] : "engine";

        // Load the plugin assembly — its module initializer must extract the
        // embedded libs and hook assembly resolution.
        var plugin = Assembly.LoadFrom(dllPath);
        var loader = plugin.GetType("IfcViewer.DependencyLoader");
        loader.GetMethod("EnsureInitialized", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public)
              .Invoke(null, null);
        string libDir = (string)loader.GetProperty("LibDir", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public)
                                      .GetValue(null);
        Console.WriteLine($"LibDir: {libDir}");
        if (libDir == null || !Directory.Exists(libDir)) { Console.WriteLine("FAIL: no lib dir"); return 1; }
        Console.WriteLine($"Extracted files: {Directory.GetFiles(libDir, "*", SearchOption.AllDirectories).Length}");

        try
        {
            return mode == "loader" ? RunPluginLoader(plugin, ifcPath) : RunGeometryPipeline(ifcPath);
        }
        catch (Exception ex)
        {
            for (var e = ex; e != null; e = e.InnerException)
                Console.WriteLine($"FAIL: {e.GetType().Name}: {e.Message}");
            return 1;
        }
    }

    // ── Mode "engine": raw xbim pipeline through the resolver ───────────────

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    static int RunGeometryPipeline(string ifcPath)
    {
        using (var model = Xbim.Ifc.IfcStore.Open(ifcPath))
        {
            Console.WriteLine($"Opened: {Path.GetFileName(ifcPath)}  instances={model.Instances.Count()}");
            var ctx = new Xbim.ModelGeometry.Scene.Xbim3DModelContext(model);
            bool ok = ctx.CreateContext(null, false);
            int shapes = ctx.ShapeInstances().Count();
            Console.WriteLine($"CreateContext={ok}  shapeInstances={shapes}");
            Console.WriteLine(ok && shapes > 0 ? "PASS" : "FAIL: no shapes");
            return ok && shapes > 0 ? 0 : 2;
        }
    }

    // ── Mode "loader": the plugin's real IfcLoader.LoadAsync (Helix scene,
    //    wall merge, opening cuts, binary cache) on a WPF dispatcher thread ──

    static int RunPluginLoader(Assembly plugin, string ifcPath)
    {
        var loaderType = plugin.GetType("IfcViewer.Ifc.IfcLoader");
        var loadAsync = loaderType.GetMethod("LoadAsync", BindingFlags.Public | BindingFlags.Static);

        int RunOnce(string tag)
        {
            object model = null;
            Exception fail = null;
            var sw = System.Diagnostics.Stopwatch.StartNew();

            var thread = new Thread(() =>
            {
                var dispatcher = Dispatcher.CurrentDispatcher;
                Task.Run(async () =>
                {
                    try
                    {
                        var task = (Task)loadAsync.Invoke(null, new object[]
                        {
                            ifcPath, dispatcher,
                            (Action<string>)(s => Console.WriteLine($"    [{tag}] {s}")),
                            default(CancellationToken)
                        });
                        await task.ConfigureAwait(false);
                        model = task.GetType().GetProperty("Result").GetValue(task);
                    }
                    catch (Exception ex) { fail = ex; }
                    finally { dispatcher.InvokeShutdown(); }
                });
                Dispatcher.Run();
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            sw.Stop();

            if (fail != null)
            {
                for (var e = fail; e != null; e = e.InnerException)
                    Console.WriteLine($"FAIL[{tag}]: {e.GetType().Name}: {e.Message}");
                return -1;
            }

            var t = model.GetType();
            int meshCount = (int)t.GetProperty("MeshCount").GetValue(model);
            int triCount = (int)t.GetProperty("TriangleCount").GetValue(model);
            var bounds = t.GetProperty("Bounds").GetValue(model);
            var handles = (System.Collections.ICollection)t.GetProperty("Handles").GetValue(model);
            Console.WriteLine($"  [{tag}] {sw.ElapsedMilliseconds} ms — meshes={meshCount} tris={triCount} handles={handles.Count} bounds={bounds}");
            return triCount;
        }

        // Force a cold load, then a cache-hit load.
        var invalidate = loaderType.GetMethod("InvalidateCache", BindingFlags.Public | BindingFlags.Static);
        invalidate.Invoke(null, new object[] { ifcPath });

        Console.WriteLine("-- cold load (full xbim pipeline + cache write)");
        int tris1 = RunOnce("cold");
        if (tris1 <= 0) { Console.WriteLine("FAIL: cold load produced no triangles"); return 2; }

        Thread.Sleep(1500); // allow async cache write to complete

        Console.WriteLine("-- warm load (binary cache hit)");
        int tris2 = RunOnce("warm");
        if (tris2 <= 0) { Console.WriteLine("FAIL: warm load produced no triangles"); return 2; }

        Console.WriteLine(tris1 == tris2
            ? $"PASS — cold and warm loads agree ({tris1} triangles)"
            : $"FAIL — cold={tris1} warm={tris2} triangle mismatch");
        return tris1 == tris2 ? 0 : 2;
    }
}
