using AlibreAddOn;
using AlibreX;
using Assimp;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
namespace AlibreAddOnAssembly
{
    public static class AlibreAddOn
    {
        private static IADRoot AlibreRoot { get; set; }
        private static IntPtr _parentWinHandle;
        private static AddOnRibbon _AssimpInsideAlibreDesignHandle;
        public static void AddOnLoad(IntPtr hwnd, IAutomationHook pAutomationHook, IntPtr unused)
        {
            AlibreRoot = (IADRoot)pAutomationHook.Root;
            _parentWinHandle = hwnd;
            _AssimpInsideAlibreDesignHandle = new AddOnRibbon(AlibreRoot, _parentWinHandle);
        }
        public static void AddOnUnload(IntPtr hwnd, bool forceUnload, ref bool cancel, int reserved1, int reserved2)
        {
            _AssimpInsideAlibreDesignHandle = null;
            AlibreRoot = null;
        }
        public static IADRoot GetRoot()
        {
            return AlibreRoot;
        }
        public static void AddOnInvoke(IntPtr hwnd, IntPtr pAutomationHook, string sessionName, bool isLicensed, int reserved1, int reserved2)
        {
            MessageBox.Show("stltostp.AddOnInvoke");
        }
        public static IAlibreAddOn GetAddOnInterface()
        {
            return _AssimpInsideAlibreDesignHandle;
        }
    }
    public class AddOnRibbon : IAlibreAddOn
    {
        private readonly MenuManager _menuManager;
        public IADRoot _AlibreRoot;
        public IntPtr _parentWinHandle;
        public AddOnRibbon(IADRoot AlibreRoot, IntPtr parentWinHandle)
        {
            _AlibreRoot = AlibreRoot;
            _parentWinHandle = parentWinHandle;
            try
            {
                _menuManager = new MenuManager(_AlibreRoot.TopmostSession);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize AddOnRibbon: {ex.Message}");
            }
        }
        public int RootMenuItem => _menuManager.GetRootMenuItem().Id;
        public IAlibreAddOnCommand? InvokeCommand(int menuID, string sessionIdentifier)
        {
            var session = _AlibreRoot.Sessions.Item(sessionIdentifier);
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem?.Command?.Invoke(session);
        }
        public bool HasSubMenus(int menuID)
        {
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem != null && menuItem.SubItems.Count > 0;
        }
        public Array? SubMenuItems(int menuID)
        {
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem?.SubItems.Select(subItem => subItem.Id).ToArray();
        }
        public string? MenuItemText(int menuID) => _menuManager.GetMenuItemById(menuID)?.Text;
        public ADDONMenuStates MenuItemState(int menuID, string sessionIdentifier) => ADDONMenuStates.ADDON_MENU_ENABLED;
        public string? MenuItemToolTip(int menuID) => _menuManager.GetMenuItemById(menuID)?.ToolTip;
        public string? MenuIcon(int menuID) => _menuManager.GetMenuItemById(menuID)?.Icon;
        public bool PopupMenu(int menuID) => false;
        public bool HasPersistentDataToSave(string sessionIdentifier) => false;
        public void SaveData(IStream pCustomData, string sessionIdentifier) { }
        public void LoadData(IStream pCustomData, string sessionIdentifier) { }
        public bool UseDedicatedRibbonTab() => false;
        private void IAlibreAddOn_setIsAddOnLicensed(bool isLicensed) { }
        void IAlibreAddOn.setIsAddOnLicensed(bool isLicensed) => IAlibreAddOn_setIsAddOnLicensed(isLicensed);
    }
    public class MenuItem
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public string ToolTip { get; set; }
        public string Icon { get; set; }
        public Func<IADSession, IAlibreAddOnCommand?>? Command { get; set; }
        public List<MenuItem> SubItems { get; set; }
        public MenuItem(int id, string text, string toolTip = "", string icon = "", Func<IADSession, IAlibreAddOnCommand?>? command = null)
        {
            Id = id;
            Text = text;
            ToolTip = toolTip;
            Icon = icon;
            Command = command;
            SubItems = new List<MenuItem>();
        }
        public void AddSubItem(MenuItem subItem) => SubItems.Add(subItem);
        public IAlibreAddOnCommand? DummyFunction(IADSession session)
        {
            MessageBox.Show($"{session.Name} : {session.FilePath}");
            return null;
        }

        public IAlibreAddOnCommand? ImportFile(IADSession session, string filterDescription, string filterExtensions, string assimpFormat = "")
        {
            var ofd = new OpenFileDialog
            {
                Title = $"Select {filterDescription} file to import",
                Filter = $"{filterDescription} files ({filterExtensions})|{filterExtensions}|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return null;

            string inputPath = ofd.FileName;
            
            try
            {
                using (var assimpService = new AssimpService())
                {
                    if (assimpService.ImportFile(inputPath, out Scene? scene) && scene != null)
                    {
                        // Convert to a format that Alibre can understand (STEP or STL)
                        string tempStepPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(inputPath)}_temp.stp");
                        
                        if (assimpService.ExportFile(scene, tempStepPath, "stp"))
                        {
                            try
                            {
                                IADRoot alibreRoot = AlibreAddOn.GetRoot();
                                if (alibreRoot != null)
                                {
                                    alibreRoot.ImportSTEPFile(tempStepPath);
                                    MessageBox.Show($"{filterDescription} file imported successfully via Assimp!",
                                        "Import Success", MessageBoxButton.OK, MessageBoxImage.Information);
                                    
                                    // Clean up temp file
                                    if (File.Exists(tempStepPath))
                                        File.Delete(tempStepPath);
                                }
                                else
                                {
                                    MessageBox.Show("Failed to get Alibre Root object. Cannot import file.",
                                                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Import failed (Alibre API error):\n{ex.Message}",
                                                "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Failed to convert {filterDescription} to STEP format for import.",
                                            "Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Failed to load {filterDescription} file with Assimp.",
                                        "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred during import:\n{ex.Message}",
                                "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            return null;
        }

        public IAlibreAddOnCommand? ExportFile(IADSession session, string filterDescription, string filterExtensions, string assimpFormat)
        {
            // Note: This is a simplified version - may need adjustment based on actual Alibre API
            var sfd = new SaveFileDialog
            {
                Title = $"Save as {filterDescription}",
                Filter = $"{filterDescription} files ({filterExtensions})|{filterExtensions}|All files (*.*)|*.*",
                DefaultExt = filterExtensions.Split(';')[0].Replace("*", "")
            };

            if (sfd.ShowDialog() != DialogResult.OK)
                return null;

            string outputPath = sfd.FileName;
            string tempStepPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(outputPath)}_export_temp.stp");

            try
            {
                // For now, show message indicating the export functionality needs current document export
                MessageBox.Show($"Export to {filterDescription} format selected.\n\n" +
                               $"Output file: {outputPath}\n\n" +
                               "Note: Full export functionality requires exporting the current document to STEP first,\n" +
                               "then converting with Assimp. This will be implemented based on the specific Alibre API methods available.",
                               "Export Function", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred during export:\n{ex.Message}",
                                "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return null;
        }
        public IAlibreAddOnCommand? Aboutmd(IADSession session)
        {
            MessageBox.Show("About stltostp - Version X.Y.Z\n(Details about the tool or link can be placed here)");
            return null;
        }

        public IAlibreAddOnCommand? AboutAssimpAddon(IADSession session)
        {
            MessageBox.Show("Assimp Inside Alibre Design Add-on\n\n" +
                           "Version 1.0.0\n" +
                           "Powered by Assimp-Net\n\n" +
                           "This add-on enables importing and exporting various 3D file formats\n" +
                           "using the powerful Assimp library.\n\n" +
                           "Supports dozens of 3D file formats including:\n" +
                           "• 3DS, OBJ, FBX, glTF, DAE (Collada)\n" +
                           "• STL, PLY, X, 3MF, and many more!\n\n" +
                           "For more information, visit:\n" +
                           "https://github.com/assimp/assimp",
                           "About Assimp Inside Alibre Design",
                           MessageBoxButton.OK,
                           MessageBoxImage.Information);
            return null;
        }

        public IAlibreAddOnCommand? ShowSupportedFormats(IADSession session)
        {
            try
            {
                using (var assimpService = new AssimpService())
                {
                    var importFormats = assimpService.GetSupportedImportFormats();
                    var exportFormats = assimpService.GetSupportedExportFormats();

                    string message = "Assimp Supported Formats:\n\n";
                    message += $"IMPORT FORMATS ({importFormats.Length}):\n";
                    message += string.Join(", ", importFormats.Take(20)) + (importFormats.Length > 20 ? "..." : "") + "\n\n";
                    message += $"EXPORT FORMATS ({exportFormats.Length}):\n";
                    message += string.Join(", ", exportFormats.Take(20)) + (exportFormats.Length > 20 ? "..." : "");

                    MessageBox.Show(message, "Supported 3D File Formats", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving format information:\n{ex.Message}",
                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return null;
        }
        public IAlibreAddOnCommand? RunCmd(IADSession session)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select file to convert to STEP",
                Filter = "3D Model files (*.stl;*.obj;*.ply;*.dae;*.3ds;*.fbx;*.gltf;*.glb)|*.stl;*.obj;*.ply;*.dae;*.3ds;*.fbx;*.gltf;*.glb|STL files (*.stl)|*.stl|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return null;

            string inputPath = ofd.FileName;
            string stepPath = Path.ChangeExtension(inputPath, ".stp");

            try
            {
                using (var assimpService = new AssimpService())
                {
                    bool success = assimpService.ConvertStlToStep(inputPath, stepPath);
                    
                    if (success && File.Exists(stepPath))
                    {
                        FileInfo stepFileInfo = new FileInfo(stepPath);
                        MessageBox.Show($"Assimp conversion completed successfully.\n\n" +
                                        $"Attempting to import STEP file from:\n{stepPath}\n\n" +
                                        $"File exists: True\n" +
                                        $"File size: {stepFileInfo.Length} bytes",
                                        "Verify STEP Path & File for Import",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Information);
                        try
                        {
                            IADRoot alibreRoot = AlibreAddOn.GetRoot();
                            if (alibreRoot != null)
                            {
                                alibreRoot.ImportSTEPFile(stepPath);
                                MessageBox.Show("STEP file imported successfully!",
                                    "Alibre Design Import Success", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            else
                            {
                                MessageBox.Show("Failed to get Alibre Root object. Cannot import STEP file.",
                                                "Alibre Design Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"STEP import failed (Alibre API error):\n{ex.Message}\n\n" +
                                            $"File was: {stepPath}",
                                            "Alibre Design Import Error",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        string errorDetails = $"Assimp failed to convert '{Path.GetFileName(inputPath)}'.\n\n";
                        if (!File.Exists(stepPath))
                            errorDetails += "The output STEP file was not created.";
                        else
                            errorDetails += "The conversion process encountered an error.";
                        
                        MessageBox.Show(errorDetails, "File Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred during the conversion process:\n{ex.Message}",
                                "Conversion Process Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return null;
        }
    }
    public class MenuManager
    {
        private readonly MenuItem _rootMenuItem;
        private readonly Dictionary<int, MenuItem> _menuItems;
        public MenuManager(IADSession topMostSession)
        {
            _rootMenuItem = new MenuItem(401, "Assimp Inside Alibre Design", "Assimp Inside Alibre Design (AIAD)");
            _menuItems = new Dictionary<int, MenuItem>();
            BuildMenus();
        }
        private void BuildMenus()
        {
            var importGroup = new MenuItem(9081, "Import");
            var import3dm = new MenuItem(9082, "3dm", "Import 3dm File", @"Icons\logo.ico");
            import3dm.Command = (session) => import3dm.ImportFile(session, "3dm", "*.3dm");
            importGroup.AddSubItem(import3dm);
            var importGltf = new MenuItem(9083, "gLTF", "Import gLTF File", @"Icons\logo.ico");
            importGltf.Command = (session) => importGltf.ImportFile(session, "glTF", "*.gltf;*.glb");
            importGroup.AddSubItem(importGltf);
            var importObj = new MenuItem(9084, "OBJ", "Import OBJ File", @"Icons\logo.ico");
            importObj.Command = (session) => importObj.ImportFile(session, "OBJ", "*.obj");
            importGroup.AddSubItem(importObj);
            var importFbx = new MenuItem(9089, "FBX", "Import FBX File", @"Icons\logo.ico");
            importFbx.Command = (session) => importFbx.ImportFile(session, "FBX", "*.fbx");
            importGroup.AddSubItem(importFbx);
            var importDae = new MenuItem(9090, "DAE", "Import DAE (Collada) File", @"Icons\logo.ico");
            importDae.Command = (session) => importDae.ImportFile(session, "DAE", "*.dae");
            importGroup.AddSubItem(importDae);
            var importPly = new MenuItem(9091, "PLY", "Import PLY File", @"Icons\logo.ico");
            importPly.Command = (session) => importPly.ImportFile(session, "PLY", "*.ply");
            importGroup.AddSubItem(importPly);
            var import3ds = new MenuItem(9092, "3DS", "Import 3DS File", @"Icons\logo.ico");
            import3ds.Command = (session) => import3ds.ImportFile(session, "3DS", "*.3ds");
            importGroup.AddSubItem(import3ds);
            var exportGroup = new MenuItem(9085, "Export");
            var export3dm = new MenuItem(9086, "3dm", "Export 3dm File", @"Icons\logo.ico");
            export3dm.Command = (session) => export3dm.ExportFile(session, "3dm", "*.3dm", "3dm");
            exportGroup.AddSubItem(export3dm);
            var exportGltf = new MenuItem(9087, "gLTF", "Export gLTF File", @"Icons\logo.ico");
            exportGltf.Command = (session) => exportGltf.ExportFile(session, "glTF", "*.gltf", "gltf2");
            exportGroup.AddSubItem(exportGltf);
            var exportDae = new MenuItem(10030, "DAE", "Export DAE (Collada) File", @"Icons\logo.ico");
            exportDae.Command = (session) => exportDae.ExportFile(session, "DAE (Collada)", "*.dae", "collada");
            exportGroup.AddSubItem(exportDae);
            var exportStl = new MenuItem(10031, "STL", "Export STL File", @"Icons\logo.ico");
            exportStl.Command = (session) => exportStl.ExportFile(session, "STL", "*.stl", "stl");
            exportGroup.AddSubItem(exportStl);
            var exportObj = new MenuItem(10032, "OBJ", "Export OBJ File", @"Icons\logo.ico");
            exportObj.Command = (session) => exportObj.ExportFile(session, "OBJ", "*.obj", "obj");
            exportGroup.AddSubItem(exportObj);
            var exportPly = new MenuItem(10033, "PLY", "Export PLY File", @"Icons\logo.ico");
            exportPly.Command = (session) => exportPly.ExportFile(session, "PLY", "*.ply", "ply");
            exportGroup.AddSubItem(exportPly);
            var exportX = new MenuItem(10034, "X", "Export X File", @"Icons\logo.ico");
            exportX.Command = (session) => exportX.ExportFile(session, "DirectX X", "*.x", "x");
            exportGroup.AddSubItem(exportX);
            var export3ds = new MenuItem(10035, "3DS", "Export 3DS File", @"Icons\logo.ico");
            export3ds.Command = (session) => export3ds.ExportFile(session, "3DS", "*.3ds", "3ds");
            exportGroup.AddSubItem(export3ds);
            var exportJson = new MenuItem(10036, "JSON", "Export Assimp JSON File", @"Icons\logo.ico");
            exportJson.Command = (session) => exportJson.ExportFile(session, "Assimp JSON", "*.json", "assjson");
            exportGroup.AddSubItem(exportJson);
            var exportAssbin = new MenuItem(10037, "ASSBIN", "Export ASSBIN File", @"Icons\logo.ico");
            exportAssbin.Command = (session) => exportAssbin.ExportFile(session, "Assimp Binary", "*.assbin", "assbin");
            exportGroup.AddSubItem(exportAssbin);
            var exportStep = new MenuItem(10038, "STEP", "Export STEP File", @"Icons\logo.ico");
            exportStep.Command = (session) => exportStep.ExportFile(session, "STEP", "*.stp;*.step", "stp");
            exportGroup.AddSubItem(exportStep);
            var exportPbrt = new MenuItem(10039, "PBRT", "Export PBRTv4 File", @"Icons\logo.ico");
            exportPbrt.Command = (session) => exportPbrt.ExportFile(session, "PBRT v4", "*.pbrt", "pbrt");
            exportGroup.AddSubItem(exportPbrt);
            var exportGltf1 = new MenuItem(10040, "gLTF 1.0", "Export glTF 1.0 File", @"Icons\logo.ico");
            exportGltf1.Command = (session) => exportGltf1.ExportFile(session, "glTF 1.0", "*.gltf", "gltf");
            exportGroup.AddSubItem(exportGltf1);
            var exportGltf2 = new MenuItem(10041, "gLTF 2.0", "Export glTF 2.0 File", @"Icons\logo.ico");
            exportGltf2.Command = (session) => exportGltf2.ExportFile(session, "glTF 2.0", "*.gltf", "gltf2");
            exportGroup.AddSubItem(exportGltf2);
            var export3mf = new MenuItem(10042, "3MF", "Export 3MF File", @"Icons\logo.ico");
            export3mf.Command = (session) => export3mf.ExportFile(session, "3MF", "*.3mf", "3mf");
            exportGroup.AddSubItem(export3mf);
            var exportFbx = new MenuItem(10043, "FBX", "Export FBX File", @"Icons\logo.ico");
            exportFbx.Command = (session) => exportFbx.ExportFile(session, "FBX", "*.fbx", "fbx");
            exportGroup.AddSubItem(exportFbx);

            var toolsGroup = new MenuItem(10050, "Tools");
            var convertTool = new MenuItem(10051, "Convert File to STEP", "Convert various 3D formats to STEP using Assimp", @"Icons\logo.ico");
            convertTool.Command = convertTool.RunCmd;
            toolsGroup.AddSubItem(convertTool);
            var aboutTool = new MenuItem(10052, "About", "About Assimp Inside Alibre Design", @"Icons\logo.ico");
            aboutTool.Command = aboutTool.AboutAssimpAddon;
            toolsGroup.AddSubItem(aboutTool);
            var formatsTool = new MenuItem(10053, "Supported Formats", "Show supported import/export formats", @"Icons\logo.ico");
            formatsTool.Command = formatsTool.ShowSupportedFormats;
            toolsGroup.AddSubItem(formatsTool);

            _rootMenuItem.AddSubItem(importGroup);
            _rootMenuItem.AddSubItem(exportGroup);
            _rootMenuItem.AddSubItem(toolsGroup);
            RegisterMenuItem(_rootMenuItem);
        }
        private void RegisterMenuItem(MenuItem menuItem)
        {
            _menuItems[menuItem.Id] = menuItem;
            foreach (var subItem in menuItem.SubItems)
                RegisterMenuItem(subItem);
        }
        public MenuItem? GetMenuItemById(int id)
        {
            _menuItems.TryGetValue(id, out MenuItem? menuItem);
            return menuItem;
        }
        public MenuItem GetRootMenuItem() => _rootMenuItem;
    }

    public class AssimpService : IDisposable
    {
        private readonly AssimpContext _assimpContext;

        public AssimpService()
        {
            _assimpContext = new AssimpContext();
        }

        public bool ConvertFile(string inputPath, string outputPath, string outputFormat)
        {
            try
            {
                // Load the input file
                var scene = _assimpContext.ImportFile(inputPath, PostProcessSteps.Triangulate | PostProcessSteps.FlipUVs);
                if (scene == null)
                {
                    return false;
                }

                // Export to the specified format
                return _assimpContext.ExportFile(scene, outputPath, outputFormat);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Assimp conversion error: {ex.Message}");
                return false;
            }
        }

        public bool ConvertStlToStep(string stlPath, string stepPath)
        {
            return ConvertFile(stlPath, stepPath, "stp");
        }

        public bool ImportFile(string filePath, out Scene? scene)
        {
            scene = null;
            try
            {
                scene = _assimpContext.ImportFile(filePath, 
                    PostProcessSteps.Triangulate | 
                    PostProcessSteps.FlipUVs | 
                    PostProcessSteps.GenerateNormals |
                    PostProcessSteps.OptimizeMeshes);
                return scene != null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Assimp import error: {ex.Message}");
                return false;
            }
        }

        public bool ExportFile(Scene scene, string outputPath, string format)
        {
            try
            {
                return _assimpContext.ExportFile(scene, outputPath, format);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Assimp export error: {ex.Message}");
                return false;
            }
        }

        public string[] GetSupportedImportFormats()
        {
            return _assimpContext.GetSupportedImportFormats();
        }

        public string[] GetSupportedExportFormats()
        {
            var exportFormats = _assimpContext.GetSupportedExportFormats();
            return exportFormats.Select(f => f.FormatId).ToArray();
        }

        public void Dispose()
        {
            _assimpContext?.Dispose();
        }
    }
}