using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Transactions;
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


    public class CustomToolStripMenuItem : UserControl
    {
        private string start_path;
        private Action<object, MouseEventArgs> on_click;
        private ToolTip toolTip;

        public void SetSelected(bool selected)
        {
            // Change the background color based on selection state
            this.BackColor = selected ? SystemColors.Highlight : SystemColors.Control;
        }

        private void CutomeMouseClick(object sender, MouseEventArgs e)
        {
            // Handle click event
            if (e.Button == MouseButtons.Right)
            {
            }
            else
            {
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = start_path,
                    UseShellExecute = true // Use this to open files with their associated applications
                });
            }

            on_click(sender, e);
        }

        private void CustomMouseLeave(object sender, EventArgs e)
        {
            // Reset background color when mouse leaves
            SetSelected(false);
            toolTip.Hide((Control)sender);
        }

        private void CustomMouseEnter(object sender, EventArgs e)
        {
            // Change background color when mouse enters
            SetSelected(true);
        }

        public CustomToolStripMenuItem(String text, Image icon, String path, Action<object , MouseEventArgs> onclick)
        {
            InitializeComponent(text, icon, path, onclick);
        }
        private void InitializeComponent(String text, Image icon, String path, Action<object, MouseEventArgs> onclick)
        {
            this.SuspendLayout();
            // 
            // CustomToolStripMenuItem
            // 
            this.Name = "CustomToolStripMenuItem";
            this.start_path = path;
            this.on_click = onclick;
            this.ResumeLayout(false);
            this.BackColor = SystemColors.Control;
            this.AutoSize = true;
            this.Padding = new Padding(4);

            var pictureBox = new PictureBox
            {
                Image = icon,
                Size = new Size(16, 16),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Margin = new Padding(0, 0, 8, 0)
            };

            var label = new Label
            {
                Text = text,
//              AutoSize = true,
                Size = new Size(150-16, 16),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            var layout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0),
                Margin = new Padding(0),
                Size = new Size(150, 16),  // 明示的に高さを固定
                AutoSize = true           // 自動サイズ調整を無効化
            };

            toolTip = new ToolTip();
            toolTip.ShowAlways = true;
            toolTip.Active = true;
            toolTip.AutomaticDelay = 500;
            toolTip.SetToolTip(label, path);
            toolTip.SetToolTip(pictureBox, path);
//          toolTip.SetToolTip(layout, path);

            layout.Controls.Add(pictureBox);
            layout.Controls.Add(label);

            this.Controls.Add(layout);
            this.MouseLeave += CustomMouseLeave;
            this.MouseEnter += CustomMouseEnter;
            this.MouseClick += CutomeMouseClick;
            label.MouseLeave += CustomMouseLeave;
            label.MouseEnter += CustomMouseEnter;
            label.MouseClick += CutomeMouseClick;
            pictureBox.MouseLeave += CustomMouseLeave;
            pictureBox.MouseEnter += CustomMouseEnter;
            pictureBox.MouseClick += CutomeMouseClick;
            layout.MouseLeave += CustomMouseLeave;
            layout.MouseEnter += CustomMouseEnter;
            label.MouseClick += CutomeMouseClick;
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
        private System.Windows.Forms.Timer closeTimer;
        private bool suppressShow = false;

        private string explorer_path;
        private string browser_path;
        private string chrome_path = "";
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
            main_menu.AutoClose = false;
            BuildMainMenuItems();

            sub_menu = new ContextMenuStrip(new Container());
            BuildSubMenuItems();

            notifyIcon = new NotifyIcon()
            {
                Icon = Resources.TTQL_main,
                Visible = true
            };

            notifyIcon.MouseMove += IconOnMouseMove;
            notifyIcon.MouseClick += Icon_OnClick;
            main_menu.ContextMenuStrip = sub_menu;
            main_menu.KeyPress += main_menu_KeyPress;

            /* MouseMoveで表示したメインメニューを、10秒でCloseする */
            closeTimer = new System.Windows.Forms.Timer
            {
                Interval = 10000, // 10秒
                Enabled = false
            };
            closeTimer.Tick += (s, e) =>
            {
                if (main_menu.Visible)
                {
                    main_menu.Close();
                }
                closeTimer.Stop();
                suppressShow = false;
            };

        }

        private void main_menu_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Escape)
            {
                System.Diagnostics.Debug.WriteLine("Escape key pressed in main menu.");
            }
            else if (e.KeyChar == (char)Keys.Enter)
            {
                System.Diagnostics.Debug.WriteLine("Enter key pressed in main menu.");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Key pressed in main menu: " + e.KeyChar);
            }
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
                            else if (assoc.ToUpper() == "CHROMEHTML")
                            {
                                // If the association is for Chrome, set the Chrome path
                                Match match = Regex.Match(command, "\"([^\"]+\\.exe)\"");
                                if (match.Success)
                                {
                                    // Extract the path to the Chrome executable
                                    command = match.Groups[1].Value;
                                    command = ResolveEnvironmentVariables(command);
                                    System.Diagnostics.Debug.WriteLine($"Chrome Path: {command}");
                                    if (string.IsNullOrEmpty(chrome_path))
                                    {
                                        chrome_path = command; // Set Chrome path
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine("No valid Chrome path found in command: " + command);
                                }
                            }
                            else if (assoc.ToUpper() == "HTTP" || assoc.ToUpper() == "HTTPS")
                            {
                                Match match = Regex.Match(command, "\"([^\"]+\\.exe)\"");
                                if (match.Success)
                                {
                                    // Extract the path to the browser executable
                                    command = match.Groups[1].Value;
                                    command = ResolveEnvironmentVariables(command);
                                    System.Diagnostics.Debug.WriteLine($"Browser Path: {command}");
                                    if (string.IsNullOrEmpty(browser_path))
                                    {
                                        browser_path = command; // Set browser path if not already set
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine("No valid browser path found in command: " + command);
                                }
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

            menu_add_path = new ToolStripMenuItem("add Path", null, AddPathShortcut);
            sub_menu.Items.Add(menu_add_path);

            menu_add_delete = new ToolStripMenuItem("delete", null, DelShortcut);
            sub_menu.Items.Add(menu_add_delete);

            menu_exit = new ToolStripMenuItem("Exit", null, Exit);
            sub_menu.Items.Add(menu_exit);
        }


        private void main_menu_onClick(object sender, MouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("main_menu_onClick : " + e.Button);
            if (e.Button == MouseButtons.Left)
            {
                main_menu.BeginInvoke(new Action(() =>
                {
                    main_menu.Close();
                }));
                
            }
            else if (e.Button == MouseButtons.Right)
            {
                
            }
        }

        private void BuildMainMenuItems()
        {
            // Clear existing items
            main_menu.Items.Clear();
            // Add shortcuts from the list
            foreach (var item in shortCutList)
            {
#if false
                ToolStripMenuItem menuItem = new ToolStripMenuItem(item.Name, item.Icon.ToBitmap(), (s, e) =>
                {
                    System.Diagnostics.Process.Start(new ProcessStartInfo
                    {
                        FileName = item.Path,
                        UseShellExecute = true // Use this to open files with their associated applications
                    });
                });
#else
                CustomToolStripMenuItem custom_item = new CustomToolStripMenuItem(item.Name, item.Icon?.ToBitmap(), item.Path, main_menu_onClick)
                {
                    Tag = item.Path,
                    AutoSize = true,
                };
                ToolStripControlHost menuItem = new ToolStripControlHost(custom_item)
                {
                    AutoSize = true,
                    Padding =  Padding.Empty,
                    Margin =  Padding.Empty,
                };
#endif
                menuItem.MouseMove += IconOnMouseMove;
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
                if (string.IsNullOrEmpty(chrome_path))
                {
                    item.Icon = Icon.ExtractAssociatedIcon(browser_path);
                }
                else
                {
                    item.Icon = Icon.ExtractAssociatedIcon(chrome_path);
                }
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
            main_menu.Show(new Point(Cursor.Position.X - 220, Cursor.Position.Y - main_menu.Height - 32));
        }

        private void DisplaySubMenu()
        {
            // Show the main menu at the current cursor position
            sub_menu.Show(new Point(Cursor.Position.X, Cursor.Position.Y - sub_menu.Height - 32));
        }

        private void Icon_OnClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (main_menu.Visible)
                {
                    System.Diagnostics.Debug.WriteLine("main_menu.Hide();");
                    main_menu.Close();
                    suppressShow = true; // Suppress showing the main menu again
                    closeTimer.Stop(); // Start the timer to close the menu after 1 second
                    closeTimer.Interval = 1000;
                    closeTimer.Start(); // Start the timer to close the menu after 1 second
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
                DisplaySubMenu();
            }
        }

        private void IconOnMouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
//          System.Diagnostics.Debug.WriteLine("MouseMove: " + e.Location);
            if (suppressShow)
            {
                return; // Do not show the main menu if suppressed
            }

            if ((main_menu.Visible == false) && (sub_menu.Visible == false))
            {
                DisplayMainMenu();
            }

            closeTimer.Stop();
            closeTimer.Interval = 10000;
            closeTimer.Start();
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

        }


        private void AddPathShortcut(object sender, EventArgs e)
        {
            BuildSubMenuItems();
//          sub_menu.Hide();
//          sub_menu.Show(sub_menu.Location);
            sub_menu.Invalidate();
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
                e.SuppressKeyPress = true;
                if (IsPathVarid(menu_path_edit.Text))
                {
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
                        string[] paths = selectedPath.Split('\\');
                        AddShortCutItem(paths[paths.Length -1], selectedPath);
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