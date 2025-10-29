using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using Microsoft.CSharp;

namespace DockableModularForm
{
    public class ModuleManifest
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }

        public static ModuleManifest FromFile(string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath) || !File.Exists(manifestPath))
            {
                return null;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                var manifest = serializer.Deserialize<ModuleManifest>(File.ReadAllText(manifestPath));
                return manifest;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ModuleManifest] Failed to parse manifest '{manifestPath}': {ex.Message}");
                return null;
            }
        }
    }

    public class ModuleDescriptor
    {
        public ModuleManifest Manifest { get; set; }
        public string DirectoryPath { get; set; }
        public string AssemblyPath { get; set; }
        public Type ModuleType { get; set; }
        public IModule Instance { get; set; }
    }

    public static class ModuleLoader
    {
        private static readonly string[] DefaultReferences = new[]
        {
            "System.dll",
            "System.Core.dll",
            "System.Drawing.dll",
            "System.Windows.Forms.dll"
        };

        public static IList<ModuleDescriptor> LoadModules(string modulesRoot)
        {
            var modules = new List<ModuleDescriptor>();

            if (string.IsNullOrWhiteSpace(modulesRoot) || !Directory.Exists(modulesRoot))
            {
                return modules;
            }

            foreach (var directory in Directory.GetDirectories(modulesRoot))
            {
                try
                {
                    var descriptor = LoadModuleFromDirectory(directory);
                    if (descriptor != null)
                    {
                        modules.Add(descriptor);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ModuleLoader] Failed to load module from '{directory}': {ex.Message}");
                }
            }

            return modules;
        }

        private static ModuleDescriptor LoadModuleFromDirectory(string directory)
        {
            var manifestPath = Path.Combine(directory, "module.json");
            var manifest = ModuleManifest.FromFile(manifestPath);

            var assemblyPath = Directory.GetFiles(directory, "*.dll", SearchOption.TopDirectoryOnly).FirstOrDefault();
            if (assemblyPath == null)
            {
                assemblyPath = TryCompileSource(directory);
            }

            if (assemblyPath == null || !File.Exists(assemblyPath))
            {
                Debug.WriteLine($"[ModuleLoader] No assembly found for module in '{directory}'.");
                return null;
            }

            Assembly assembly;
            try
            {
                assembly = Assembly.LoadFrom(assemblyPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ModuleLoader] Failed to load assembly '{assemblyPath}': {ex.Message}");
                return null;
            }

            var moduleType = assembly.GetTypes()
                .FirstOrDefault(t => typeof(IModule).IsAssignableFrom(t) && !t.IsAbstract && t.IsClass);

            if (moduleType == null)
            {
                Debug.WriteLine($"[ModuleLoader] No IModule implementation found in '{assembly.FullName}'.");
                return null;
            }

            IModule instance;
            try
            {
                instance = (IModule)Activator.CreateInstance(moduleType);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ModuleLoader] Failed to instantiate module '{moduleType.FullName}': {ex.Message}");
                return null;
            }

            return new ModuleDescriptor
            {
                Manifest = manifest,
                DirectoryPath = directory,
                AssemblyPath = assemblyPath,
                ModuleType = moduleType,
                Instance = instance
            };
        }

        private static string TryCompileSource(string directory)
        {
            var sourcePath = Directory.GetFiles(directory, "*.cs", SearchOption.TopDirectoryOnly).FirstOrDefault();
            if (sourcePath == null)
            {
                return null;
            }

            var provider = new CSharpCodeProvider();
            var outputAssembly = Path.Combine(directory, Path.GetFileNameWithoutExtension(sourcePath) + ".dll");

            var parameters = new CompilerParameters
            {
                GenerateExecutable = false,
                OutputAssembly = outputAssembly,
                GenerateInMemory = false,
                IncludeDebugInformation = false
            };

            foreach (var reference in DefaultReferences)
            {
                parameters.ReferencedAssemblies.Add(reference);
            }

            parameters.ReferencedAssemblies.Add("System.Web.Extensions.dll");
            parameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");

            CompilerResults results = provider.CompileAssemblyFromFile(parameters, sourcePath);
            if (results.Errors.HasErrors)
            {
                var errors = string.Join(Environment.NewLine, results.Errors.Cast<CompilerError>().Select(e => e.ToString()));
                Debug.WriteLine($"[ModuleLoader] Failed to compile module source '{sourcePath}': {errors}");
                return null;
            }

            var assemblyPath = !string.IsNullOrEmpty(results.PathToAssembly) ? results.PathToAssembly : outputAssembly;
            return File.Exists(assemblyPath) ? assemblyPath : null;
        }
    }
}
