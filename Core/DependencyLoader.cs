using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

#if !NETFRAMEWORK
using System.Runtime.Loader;
#endif

namespace IfcViewer
{
    /// <summary>
    /// Makes IfcViewer.dll fully self-contained.
    ///
    /// At build time every dependency DLL (managed NuGet assemblies plus the
    /// native xBIM/OpenCascade geometry engine under <c>win-x64\</c>) is
    /// gzip-compressed and embedded as a resource named
    /// <c>IfcViewer.Libs/&lt;relative-path&gt;.gz</c> (see EmbedDependencyLibs
    /// in IfcViewer.csproj).
    ///
    /// At runtime <see cref="EnsureInitialized"/>:
    ///   1. extracts those resources (once per build, keyed by the assembly
    ///      MVID) to <c>%LOCALAPPDATA%\RKTools\IfcViewer\libs\&lt;tfm&gt;\&lt;mvid&gt;</c>,
    ///   2. hooks AssemblyResolve (and AssemblyLoadContext.Resolving on .NET 8)
    ///      to load managed assemblies from that folder by simple name,
    ///   3. registers the folder (and its win-x64 subfolder) on the native DLL
    ///      search path so the geometry engine and OpenCascade DLLs resolve.
    ///
    /// A module initializer guarantees this runs before any other code in this
    /// assembly executes — i.e. before the JIT ever needs a dependency type.
    /// </summary>
    internal static class DependencyLoader
    {
        private const string ResourcePrefix = "IfcViewer.Libs/";

        private static readonly object Gate = new object();
        private static bool _initialized;

        /// <summary>Directory all dependencies were extracted to (null until initialized).</summary>
        internal static string LibDir { get; private set; }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr AddDllDirectory(string lpPathName);

        [System.Runtime.CompilerServices.ModuleInitializer]
        internal static void EnsureInitialized()
        {
            lock (Gate)
            {
                if (_initialized) return;
                _initialized = true;

                try
                {
                    Initialize();
                }
                catch (Exception ex)
                {
                    // Never let dependency extraction take down Revit's addin load.
                    // The resolver below still probes the assembly's own folder, so a
                    // loose-file deployment keeps working even if extraction failed.
                    try { SessionLogger.Error("DependencyLoader failed to initialize.", ex); } catch { }
                }
            }
        }

        private static void Initialize()
        {
            var assembly = typeof(DependencyLoader).Assembly;

            // Resolver first — it is useful even when there is nothing to extract
            // (loose-file builds with -p:EmbedDependencies=false).
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
#if !NETFRAMEWORK
            AssemblyLoadContext.GetLoadContext(assembly)!.Resolving += OnAlcResolving;
#endif

            string[] resourceNames = assembly.GetManifestResourceNames()
                .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal))
                .ToArray();
            if (resourceNames.Length == 0)
            {
                // Loose-file build: dependencies (incl. the native engine under
                // win-x64\) sit next to IfcViewer.dll — register those paths instead.
                string selfDir = null;
                try { selfDir = Path.GetDirectoryName(assembly.Location); } catch { }
                if (!string.IsNullOrEmpty(selfDir))
                    RegisterNativeSearchPaths(selfDir);
                SessionLogger.Info("DependencyLoader: no embedded libs (loose-file build).");
                return;
            }

            // One extraction folder per build: the MVID changes on every compile,
            // so a new build never collides with files locked by another Revit
            // session running an older build.
            string mvid = assembly.ManifestModule.ModuleVersionId.ToString("N").Substring(0, 16);
            string tfm =
#if NETFRAMEWORK
                "net48";
#else
                "net8.0";
#endif
            string root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RKTools", "IfcViewer", "libs", tfm);
            LibDir = Path.Combine(root, mvid);

            var sw = System.Diagnostics.Stopwatch.StartNew();
            string marker = Path.Combine(LibDir, ".complete");
            if (!File.Exists(marker))
            {
                Directory.CreateDirectory(LibDir);
                foreach (string resName in resourceNames)
                {
                    // "IfcViewer.Libs/win-x64/TKernel.dll.gz" → "win-x64\TKernel.dll"
                    string rel = resName.Substring(ResourcePrefix.Length);
                    if (rel.EndsWith(".gz", StringComparison.OrdinalIgnoreCase))
                        rel = rel.Substring(0, rel.Length - 3);
                    string dest = Path.Combine(LibDir, rel.Replace('/', Path.DirectorySeparatorChar));
                    Directory.CreateDirectory(Path.GetDirectoryName(dest));

                    using (var res = assembly.GetManifestResourceStream(resName))
                    using (var gz = new GZipStream(res, CompressionMode.Decompress))
                    using (var file = File.Create(dest))
                        gz.CopyTo(file);
                }
                File.WriteAllText(marker, DateTime.UtcNow.ToString("o"));
                CleanupOldExtractions(root, keep: LibDir);
            }
            sw.Stop();

            RegisterNativeSearchPaths(LibDir);

            SessionLogger.Info(
                $"DependencyLoader: {resourceNames.Length} embedded libs ready in {sw.ElapsedMilliseconds} ms → {LibDir}");
        }

        /// <summary>
        /// Registers <paramref name="baseDir"/> (and its win-x64 subfolder — home
        /// of the xBIM geometry engine and the OpenCascade TK*.dll set) on the
        /// native DLL search path.
        /// </summary>
        private static void RegisterNativeSearchPaths(string baseDir)
        {
            string nativeDir = Path.Combine(baseDir, "win-x64");
            SetDllDirectory(Directory.Exists(nativeDir) ? nativeDir : baseDir);
            AddDllDirectory(baseDir);
            if (Directory.Exists(nativeDir)) AddDllDirectory(nativeDir);
            PrependToPath(baseDir);
            if (Directory.Exists(nativeDir)) PrependToPath(nativeDir);
        }

        private static void PrependToPath(string dir)
        {
            string path = Environment.GetEnvironmentVariable("PATH") ?? "";
            if (path.IndexOf(dir, StringComparison.OrdinalIgnoreCase) < 0)
                Environment.SetEnvironmentVariable("PATH", dir + ";" + path);
        }

        /// <summary>Best-effort removal of extraction folders left by older builds.</summary>
        private static void CleanupOldExtractions(string root, string keep)
        {
            try
            {
                foreach (string dir in Directory.GetDirectories(root))
                {
                    if (string.Equals(dir, keep, StringComparison.OrdinalIgnoreCase)) continue;
                    try { Directory.Delete(dir, recursive: true); }
                    catch { /* locked by another Revit session — leave it */ }
                }
            }
            catch { }
        }

        // ── Managed assembly resolution ──────────────────────────────────────

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
            => ResolveByName(args.Name);

#if !NETFRAMEWORK
        private static Assembly OnAlcResolving(AssemblyLoadContext context, AssemblyName name)
        {
            string path = FindAssemblyFile(name.Name);
            return path != null ? context.LoadFromAssemblyPath(path) : null;
        }
#endif

        private static Assembly ResolveByName(string fullName)
        {
            try
            {
                string simpleName = new AssemblyName(fullName).Name;
                if (simpleName == null || simpleName.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
                    return null;

                // Already loaded (any version)? Return it — mirrors a binding redirect.
                foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (!loaded.IsDynamic
                        && string.Equals(loaded.GetName().Name, simpleName, StringComparison.OrdinalIgnoreCase))
                        return loaded;
                }

                string path = FindAssemblyFile(simpleName);
                if (path == null)
                {
                    SessionLogger.Warn($"AssemblyResolve: '{simpleName}' not found.");
                    return null;
                }
                return Assembly.LoadFrom(path);
            }
            catch (Exception ex)
            {
                SessionLogger.Error($"AssemblyResolve failed for '{fullName}'", ex);
                return null;
            }
        }

        /// <summary>
        /// Probes the extraction folder (and its immediate subfolders), then the
        /// folder IfcViewer.dll itself was loaded from, for &lt;name&gt;.dll.
        /// </summary>
        private static string FindAssemblyFile(string simpleName)
        {
            foreach (string baseDir in CandidateDirs())
            {
                string candidate = Path.Combine(baseDir, simpleName + ".dll");
                if (File.Exists(candidate)) return candidate;

                try
                {
                    foreach (string sub in Directory.GetDirectories(baseDir))
                    {
                        candidate = Path.Combine(sub, simpleName + ".dll");
                        if (File.Exists(candidate)) return candidate;
                    }
                }
                catch { }
            }
            return null;
        }

        private static IEnumerable<string> CandidateDirs()
        {
            if (LibDir != null && Directory.Exists(LibDir))
                yield return LibDir;

            string selfDir = null;
            try { selfDir = Path.GetDirectoryName(typeof(DependencyLoader).Assembly.Location); }
            catch { }
            if (!string.IsNullOrEmpty(selfDir) && Directory.Exists(selfDir))
                yield return selfDir;
        }
    }
}

#if NETFRAMEWORK
namespace System.Runtime.CompilerServices
{
    /// <summary>Enables C# module initializers when targeting .NET Framework.</summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    internal sealed class ModuleInitializerAttribute : Attribute { }
}
#endif
