using Microsoft.Win32;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Resources;
using System.Text;
using System.Windows.Forms;
using TaskTrayQuickLaunch.Properties;


namespace TaskTrayQuickLaunch
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // terminate runcat if there's any existing instance
            var procMutex = new System.Threading.Mutex(true, "_TTQL_MUTEX", out var result);
            if (!result)
            {
                return;
            }

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.Run(new TTQLApplicationContext());

            procMutex.ReleaseMutex();
        }
    }

    public class TTQLApplicationContext : ApplicationContext
    {
        private const String INI_FILE_NAME = "TaskTrayQuickLaunch.ini";
        private NotifyIcon notifyIcon;
        private ContextMenuStrip sub_menu;
        private ContextMenuStrip main_menu;
        private ToolStripMenuItem menu_add_file;
        private ToolStripMenuItem menu_add_folder;
        private ToolStripMenuItem menu_add_path;
        private ToolStripMenuItem menu_add_delete;
        private ToolStripMenuItem menu_exit;
        private ToolStripTextBox menu_path_edit;
        private Point menu_position;

        private bool sub_menu_add_mode;
        private string explorer_path;
        private string browser_path;
        private List<ShortCutItem> shortCutList = new List<ShortCutItem>();

        private struct ShortCutItem
        {
            public string Name;
            public string Path;
            public Icon Icon;
        }

        public TTQLApplicationContext()
        {
            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            GetFtypeText();

            LoadShortCutIni();
            main_menu = new ContextMenuStrip(new Container());
            main_menu.Closing += main_menu_OnClose;
            BuildMainMenuItems();

            sub_menu_add_mode = false;
            sub_menu = new ContextMenuStrip(new Container());
            BuildSubMenuItems();

            notifyIcon = new NotifyIcon()
            {
                Icon = Resources.TTQL_main,
                Visible = true
            };

            notifyIcon.MouseMove += IconOnMouseMove;
            notifyIcon.MouseClick += Icon_OnClick;
//          notifyIcon.ContextMenuStrip = sub_menu;
        }

        private string ResolveEnvironmentVariables(string path)
        {
            // Resolve environment variables in paths
            if (!string.IsNullOrEmpty(path))
            {
                path = Environment.ExpandEnvironmentVariables(path);
            }
            return path;
        }


        /* Ftypeコマンドで得られる情報から、エクスプローラとブラウザのEXEパスを取得する */
        private void GetFtypeText()
        {
            try
            {
                var p = System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c ftype",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
                var output = p.StandardOutput.ReadToEnd();
                foreach (var line in output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (line.Contains("="))
                    {
                        var parts = line.Split('=');
                        if (parts.Length == 2)
                        {
                            var assoc = parts[0].Trim();
                            var command = parts[1].Trim();
                            // Here you can process the association, e.g., log it or store it
                            System.Diagnostics.Debug.WriteLine($"Association: {assoc}, Command: {command}");
                            if (assoc.ToUpper() == "FOLDER")
                            {
                                // If the association is for folders, set the explorer path
                                explorer_path = ResolveEnvironmentVariables(command);
                            }
                            else if (assoc.ToUpper() == "HTTP" || assoc.ToUpper() == "HTTPS")
                            {
                                // If the association is for HTTP/HTTPS, set the browser path
                                browser_path = command;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Exception: " + e.Message);
            }
        }

        private void BuildSubMenuItems()
        {
            // Clear existing items
            sub_menu.Items.Clear();
            // Add shortcuts from the list
            sub_menu.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripMenuItem($"{Application.ProductName} v{Application.ProductVersion}")
                {
                    Enabled = false
                },
                new ToolStripSeparator()
            });

            menu_add_file = new ToolStripMenuItem("add File", null, AddFileShortcut);
            sub_menu.Items.Add(menu_add_file);

            menu_add_folder = new ToolStripMenuItem("add Folder", null, AddFolderShortcut);
            sub_menu.Items.Add(menu_add_folder);

            if (sub_menu_add_mode)
            {
                menu_path_edit = new ToolStripTextBox();
                menu_path_edit.Text = "Path or URL";
                menu_path_edit.ToolTipText = "Input File / Folder path or URL";
                menu_path_edit.KeyDown += OnKeyDown;
                sub_menu.Items.Add(menu_path_edit);
            }
            else
            {
                menu_add_path = new ToolStripMenuItem("add Path", null, AddPathShortcut);
                sub_menu.Items.Add(menu_add_path);
            }

            menu_add_delete = new ToolStripMenuItem("delete", null, DelShortcut);
            sub_menu.Items.Add(menu_add_delete);

            menu_exit = new ToolStripMenuItem("Exit", null, Exit);
            sub_menu.Items.Add(menu_exit);
        }

        private void BuildMainMenuItems()
        {
            // Clear existing items
            main_menu.Items.Clear();
            // Add shortcuts from the list
            foreach (var item in shortCutList)
            {
                ToolStripMenuItem menuItem = new ToolStripMenuItem(item.Name, item.Icon.ToBitmap(), (s, e) =>
                {
                    System.Diagnostics.Process.Start(new ProcessStartInfo
                    {
                        FileName = item.Path,
                        UseShellExecute = true // Use this to open files with their associated applications
                    });
                });
                main_menu.Items.Add(menuItem);
            }
        }


        private void AddShortCutItem(string name, string path)
        {
            foreach (var tmp_item in shortCutList)
            {
                if (tmp_item.Path.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    // Item already exists, do not add again
                    return;
                }
            }

            ShortCutItem item = new ShortCutItem
            {
                Name = name,
                Path = path,
                Icon = null // Default icon, will be set later if file exists
            };
            if (File.Exists(path))
            {
                item.Icon = Icon.ExtractAssociatedIcon(path);
            }
            else if (Directory.Exists(path))
            {
                // If it's a directory, use a folder icon
                item.Icon = Icon.ExtractAssociatedIcon(explorer_path);
            }
            else if (Uri.IsWellFormedUriString(path, UriKind.Absolute))
            {
                // If it's a URL, use a browser icon
                item.Icon = Icon.ExtractAssociatedIcon(browser_path);
            }
            else
            {
                item.Icon = null; // No icon if file doesn't exist
            }

            shortCutList.Add(item);
        }

        private void LoadShortCutIni()
        {
            try
            {
                StreamReader sr = new StreamReader(INI_FILE_NAME);
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line))
                    {
                        continue; // Skip comments and empty lines
                    }
                    var parts = line.Split('|');
                    if (parts.Length == 2)
                    {
                        AddShortCutItem(parts[0].Trim(), parts[1].Trim());
                    }
                }

                sr.Close();
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Exception: " + e.Message);
                SaveShortCutIni();
            }
            finally
            {
                System.Diagnostics.Debug.WriteLine("Executing finally block.");
            }
        }

        private void SaveShortCutIni()
        {
            try
            {
                StreamWriter sw = new StreamWriter(INI_FILE_NAME);
                sw.WriteLine("# TaskTrayQuickLaunch v" + Application.ProductVersion);
                foreach (var item in shortCutList)
                {
                    sw.WriteLine($"{item.Name}|{item.Path}");
                }

                sw.Close();
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                System.Diagnostics.Debug.WriteLine("Executing finally block.");
            }
        }

        private void DisplayMainMenu()
        {
            // Show the main menu at the current cursor position
            main_menu.Show(new Point(Cursor.Position.X, Cursor.Position.Y - main_menu.Height - 32));
        }

        private void Icon_OnClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (main_menu.Visible)
                {
                    main_menu.Hide();
                }
                else
                {
                    // Show the main menu at the current cursor position
                    DisplayMainMenu();
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                if (main_menu.Visible)
                {
                    main_menu.Hide();
                }
                sub_menu.Show(Cursor.Position);
            }
        }

        private void IconOnMouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("MouseMove: " + e.Location);
            if ((main_menu.Visible == false) && (sub_menu.Visible == false))
            {
                DisplayMainMenu();
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

        }

        private void main_menu_OnClose(object sender, ToolStripDropDownClosingEventArgs e)
        {

        }

        private void AddPathShortcut(object sender, EventArgs e)
        {
            sub_menu_add_mode = true;
            BuildSubMenuItems();
            sub_menu.Show(sub_menu.Location);
        }

        private bool IsPathVarid(string path)
        {
            if (File.Exists(path) || Directory.Exists(path))
            {
                return true; // Valid path
            }
            else if (Uri.IsWellFormedUriString(path, UriKind.Absolute))
            {
                return true; // Valid URL
            }   

            return false; // Invalid path or URL
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                System.Diagnostics.Debug.WriteLine("OnPathEnter : " + menu_path_edit.Text);
                sub_menu_add_mode = false;
                //              e.Handled = true;
                e.SuppressKeyPress = true;
                if (IsPathVarid(menu_path_edit.Text))
                {
                    sub_menu_add_mode = false;
                    AddShortCutItem(Path.GetFileNameWithoutExtension(menu_path_edit.Text), menu_path_edit.Text);
                    sub_menu.Hide();
                }
                else
                {

                }
            }
        }

        private void AddFolderShortcut(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a folder to add as a shortcut";
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    if (Directory.Exists(selectedPath))
                    {
                        AddShortCutItem(Path.GetFileNameWithoutExtension(selectedPath), selectedPath);
                        SaveShortCutIni();
                    }
                }
            }
        }

        private void AddFileShortcut(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "File|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                if (File.Exists(filePath))
                {
                    // Add the file to the shortcut list
                    // For now, just log it to debug output
                    System.Diagnostics.Debug.WriteLine("Added shortcut: " + filePath);
                    AddShortCutItem(Path.GetFileNameWithoutExtension(filePath), filePath);
                    SaveShortCutIni();
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("No file selected.");
            }
        }

        private void DelShortcut(object sender, EventArgs e)
        {

        }

        private void Exit(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }
    }

}