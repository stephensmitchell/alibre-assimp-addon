using AlibreAddOn;
using AlibreX;
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
        public IAlibreAddOnCommand? Aboutmd(IADSession session)
        {
            MessageBox.Show("About stltostp - Version X.Y.Z\n(Details about the tool or link can be placed here)");
            return null;
        }
        public IAlibreAddOnCommand? RunCmd(IADSession session)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select STL file to convert",
                Filter = "STL files (*.stl)|*.stl",
                CheckFileExists = true,
                Multiselect = false
            };
            if (ofd.ShowDialog() != DialogResult.OK)
                return null;
            string stlPath = ofd.FileName;
            string stepPath = Path.ChangeExtension(stlPath, ".stp");
            string addInDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
            string exePath = Path.Combine(addInDir, "tools", "stltostp.exe");
            if (!File.Exists(exePath))
            {
                MessageBox.Show($"Converter not found at expected location:\n{exePath}",
                                "stltostp.exe missing", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            string args = $"\"{stlPath}\" \"{stepPath}\"";
            try
            {
                var psi = new ProcessStartInfo(exePath, args)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var proc = Process.Start(psi);
                if (proc == null)
                {
                    MessageBox.Show("Failed to launch stltostp.exe. The process could not be started.",
                                    "STL → STEP Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }
                string stdOutput = proc.StandardOutput.ReadToEnd();
                string stdError = proc.StandardError.ReadToEnd();
                proc.WaitForExit();
                if (proc.ExitCode == 0 && File.Exists(stepPath))
                {
                    FileInfo stepFileInfo = new FileInfo(stepPath);
                    MessageBox.Show($"stltostp.exe completed successfully.\n\n" +
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
                            MessageBox.Show(alibreRoot.AppTitle,
                                "Alibre Design Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            ;
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
                    string errorDetails = $"stltostp.exe failed to convert '{Path.GetFileName(stlPath)}'.\n\n";
                    errorDetails += $"Exit Code: {proc.ExitCode}\n";
                    if (!string.IsNullOrWhiteSpace(stdOutput)) 
                        errorDetails += $"\nStandard Output:\n{stdOutput}\n";
                    if (!string.IsNullOrWhiteSpace(stdError))
                        errorDetails += $"\nStandard Error:\n{stdError}\n";
                    if (!File.Exists(stepPath))
                        errorDetails += "\nThe output STEP file was not found at the expected location:\n" + stepPath;
                    else if (proc.ExitCode != 0)
                        errorDetails += "\nThe tool exited with an error. The STEP file might be incomplete or corrupt even if it exists.";
                    MessageBox.Show(errorDetails, "STL → STEP Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred during the STL to STEP process:\n{ex.Message}",
                                "STL → STEP Process Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            import3dm.Command = import3dm.DummyFunction;
            importGroup.AddSubItem(import3dm);
            var importGltf = new MenuItem(9083, "gLTF", "Import gLTF File", @"Icons\logo.ico");
            importGltf.Command = importGltf.DummyFunction;
            importGroup.AddSubItem(importGltf);
            var exportGroup = new MenuItem(9085, "Export");
            var export3dm = new MenuItem(9086, "3dm", "Export 3dm File", @"Icons\logo.ico");
            export3dm.Command = export3dm.DummyFunction;
            exportGroup.AddSubItem(export3dm);
            var exportGltf = new MenuItem(9087, "gLTF", "Export gLTF File", @"Icons\logo.ico");
            exportGltf.Command = exportGltf.DummyFunction;
            exportGroup.AddSubItem(exportGltf);
            var exportDae = new MenuItem(10030, "DAE", "Export DAE (Collada) File", @"Icons\logo.ico");
            exportDae.Command = exportDae.DummyFunction;
            exportGroup.AddSubItem(exportDae);
            var exportStl = new MenuItem(10031, "STL", "Export STL File", @"Icons\logo.ico");
            exportStl.Command = exportStl.DummyFunction;
            exportGroup.AddSubItem(exportStl);
            var exportObj = new MenuItem(10032, "OBJ", "Export OBJ File", @"Icons\logo.ico");
            exportObj.Command = exportObj.DummyFunction;
            exportGroup.AddSubItem(exportObj);
            var exportPly = new MenuItem(10033, "PLY", "Export PLY File", @"Icons\logo.ico");
            exportPly.Command = exportPly.DummyFunction;
            exportGroup.AddSubItem(exportPly);
            var exportX = new MenuItem(10034, "X", "Export X File", @"Icons\logo.ico");
            exportX.Command = exportX.DummyFunction;
            exportGroup.AddSubItem(exportX);
            var export3ds = new MenuItem(10035, "3DS", "Export 3DS File", @"Icons\logo.ico");
            export3ds.Command = export3ds.DummyFunction;
            exportGroup.AddSubItem(export3ds);
            var exportJson = new MenuItem(10036, "JSON", "Export Assimp JSON File", @"Icons\logo.ico");
            exportJson.Command = exportJson.DummyFunction;
            exportGroup.AddSubItem(exportJson);
            var exportAssbin = new MenuItem(10037, "ASSBIN", "Export ASSBIN File", @"Icons\logo.ico");
            exportAssbin.Command = exportAssbin.DummyFunction;
            exportGroup.AddSubItem(exportAssbin);
            var exportStep = new MenuItem(10038, "STEP", "Export STEP File", @"Icons\logo.ico");
            exportStep.Command = exportStep.DummyFunction;
            exportGroup.AddSubItem(exportStep);
            var exportPbrt = new MenuItem(10039, "PBRT", "Export PBRTv4 File", @"Icons\logo.ico");
            exportPbrt.Command = exportPbrt.DummyFunction;
            exportGroup.AddSubItem(exportPbrt);
            var exportGltf1 = new MenuItem(10040, "gLTF 1.0", "Export glTF 1.0 File", @"Icons\logo.ico");
            exportGltf1.Command = exportGltf1.DummyFunction;
            exportGroup.AddSubItem(exportGltf1);
            var exportGltf2 = new MenuItem(10041, "gLTF 2.0", "Export glTF 2.0 File", @"Icons\logo.ico");
            exportGltf2.Command = exportGltf2.DummyFunction;
            exportGroup.AddSubItem(exportGltf2);
            var export3mf = new MenuItem(10042, "3MF", "Export 3MF File", @"Icons\logo.ico");
            export3mf.Command = export3mf.DummyFunction;
            exportGroup.AddSubItem(export3mf);
            var exportFbx = new MenuItem(10043, "FBX", "Export FBX File", @"Icons\logo.ico");
            exportFbx.Command = exportFbx.DummyFunction;
            exportGroup.AddSubItem(exportFbx);
            _rootMenuItem.AddSubItem(importGroup);
            _rootMenuItem.AddSubItem(exportGroup);
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
}