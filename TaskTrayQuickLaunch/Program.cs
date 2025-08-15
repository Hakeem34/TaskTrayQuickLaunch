#define VALID_CODE

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Sockets;
using System.Reflection;
using System.Reflection.Emit;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
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


    /************************************************************************************************/
    /* 定数定義                                                                                     */
    /************************************************************************************************/
    static class Constants
    {
        public const int SHORTCUT_WIDTH         = 200;
        public const int ICON_SIZE              = 16;
        public const int LABEL_HEIGHT           = 16;
        public const int TEXT_BOX_HEIGHT        = 24;
        public const int MENU_HIDE_INTERVAL     = 10000;
        public const int MENU_SUPPRESS_INTERVAL = 1000;
    }


    /************************************************************************************************/
    /* カスタムアイテムからアプリケーションへイベント通知用インターフェイス                         */
    /************************************************************************************************/
    public interface ApplicationEventHandler
    {
        public void item_on_click(object sender, MouseEventArgs e);
        public void item_add_path(string name, string path);
        public void on_focus(int index);
    }


    /************************************************************************************************/
    /* メインメニュー用のカスタムアイテムクラス                                                     */
    /************************************************************************************************/
    public class CustomToolStripMenuItem : UserControl
    {
        private string start_path;
        private ToolTip toolTip;
        private bool locked = false;
        private ApplicationEventHandler event_handler;

        /* コンストラクタ */
        public CustomToolStripMenuItem(String text, Image icon, String path, ApplicationEventHandler event_handler)
        {
            InitializeComponent(text, icon, path);

            /* 上位からイベントハンドラを登録しておく */
            this.event_handler = event_handler;
        }

        /* 初期化処理 */
        private void InitializeComponent(String text, Image icon, String path)
        {
            this.SuspendLayout();
            // 
            // CustomToolStripMenuItem
            // 
            this.Name = "CustomToolStripMenuItem";
            this.start_path = path;
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
            pictureBox.MouseLeave += CustomMouseLeave;
            pictureBox.MouseEnter += CustomMouseEnter;
            pictureBox.MouseClick += CustomMouseClick;

            var layout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0),
                Margin = new Padding(0),
                AutoSize = true           // 自動サイズ調整を無効化
            };
            layout.MouseLeave += CustomMouseLeave;
            layout.MouseEnter += CustomMouseEnter;
            layout.MouseClick += CustomMouseClick;
            layout.Controls.Add(pictureBox);

            toolTip = new ToolTip();
            toolTip.ShowAlways = true;
            toolTip.Active = true;
            toolTip.AutomaticDelay = 500;
            toolTip.SetToolTip(pictureBox, path);
            if (text != "")
            {
                var label = new System.Windows.Forms.Label
                {
                    Text = text,
                    Size = new Size(150 - Constants.ICON_SIZE, Constants.LABEL_HEIGHT),
                    Dock = DockStyle.Fill
                };
                label.MouseLeave += CustomMouseLeave;
                label.MouseEnter += CustomMouseEnter;
                label.MouseClick += CustomMouseClick;

                toolTip.SetToolTip(label, path);
                layout.Size = new Size(Constants.SHORTCUT_WIDTH, Constants.LABEL_HEIGHT);
                layout.Controls.Add(label);
            }
            else
            {
                var text_box = new TextBox
                {
                    Text = "",
                    Size = new Size(150 - Constants.ICON_SIZE, Constants.TEXT_BOX_HEIGHT),
                    Dock = DockStyle.Fill
                };

                /* クリップボードにテキストが入っている場合は、それを初期値にする */
                if (Clipboard.ContainsText())
                {
                    string clip_text = Clipboard.GetText();
                    Console.WriteLine("Clipboard Text: " + clip_text);
                    text_box.Text = clip_text;
                }
                else
                {
                    text_box.Text = "new shortcut";
                }

                    text_box.KeyDown += CustomKeyDown;
                layout.Size = new Size(Constants.SHORTCUT_WIDTH, Constants.TEXT_BOX_HEIGHT);
                layout.Controls.Add(text_box);
            }

            this.Controls.Add(layout);
            this.MouseLeave += CustomMouseLeave;
            this.MouseEnter += CustomMouseEnter;
            this.MouseClick += CustomMouseClick;
        }

        public string GetStartPath()
        {
            // Return the start path associated with this item
            return start_path;
        }

        /* アイテムの選択状態 取得 */
        public bool IsSelected()
        {   // Check if the background color is the highlight color
            return this.BackColor == SystemColors.Highlight;
        }

        /* アイテムの選択状態 固定 */
        public void SetSelectLock(bool locked)
        {
            this.locked = locked;
        }

        /* アイテムの選択状態 設定 */
        public void SetSelected(bool selected)
        {
            // Change the background color based on selection state
            if (locked)
            {
                return; // If locked, do not change the selection state
            }

            /* メニュー上で操作がある場合は、タイムアウトを延長する */
            if (selected)
            {
                event_handler.on_focus((int)this.Tag);
            }
            this.BackColor = selected ? SystemColors.Highlight : SystemColors.Control;
        }

        /* TextBoxへの入力確定処理 */
        private void CustomKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (sender is TextBox text_box)
                {
                    System.Diagnostics.Debug.WriteLine("Enter Key Pressed! : " + text_box.Text);
                    if (Uri.IsWellFormedUriString(text_box.Text, UriKind.Absolute))
                    {
                        string[] paths = text_box.Text.Split('/');
                        string name = paths[paths.Length - 1];
                        if (name == "")
                        {
                            if (paths.Length > 1)
                            {
                                name = paths[paths.Length - 2];
                            }
                            else
                            {
                                name = text_box.Text;
                            }
                        }
                        event_handler.item_add_path(name, text_box.Text);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                    else if (File.Exists(text_box.Text))
                    {
                        event_handler.item_add_path(Path.GetFileNameWithoutExtension(text_box.Text), text_box.Text);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                    else if (Directory.Exists(text_box.Text))
                    {
                        string[] paths = text_box.Text.Split('\\');
                        event_handler.item_add_path(paths[paths.Length -1], text_box.Text);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                }
            }
        }

        /* アイテムクリック時のハンドラ */
        private void CustomMouseClick(object sender, MouseEventArgs e)
        {
            // Handle click event
            if (e.Button == MouseButtons.Right)
            {
            }
            else if (start_path != "")
            {
                /* シェルからstart経由でショートカットを呼び出す */
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = start_path,
                    UseShellExecute = true // Use this to open files with their associated applications
                });
            }

            event_handler.item_on_click(sender, e);
        }

        /* マウスカーソルがアイテムから外れた */
        private void CustomMouseLeave(object sender, EventArgs e)
        {
            // Reset background color when mouse leaves
            SetSelected(false);
            toolTip.Hide((Control)sender);
        }

        /* マウスカーソルがアイテムに入った */
        private void CustomMouseEnter(object sender, EventArgs e)
        {
            // Change background color when mouse enters
            SetSelected(true);
        }
    }


    public class CustomMenuStrip : ContextMenuStrip
    {
        public CustomMenuStrip(Container container)
            : base(container) // 明示してもしなくてもOK
        {
            // 初期化コード
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            System.Diagnostics.Debug.WriteLine("main_nenu.keydown! : " + keyData);
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true; // ハンドリング済み
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }

    /************************************************************************************************/
    /* ショートカットアイテム情報保持の構造体                                                       */
    /************************************************************************************************/
    public struct ShortCutItem
    {
        public string Name;
        public string Path;
        public string Group;
        public Icon Icon;
    }

    /************************************************************************************************/
    /* アプリケーションのメインクラス                                                               */
    /************************************************************************************************/
    public class TTQLApplicationContext : ApplicationContext, ApplicationEventHandler
    {
        private const String INI_FILE_NAME = "TaskTrayQuickLaunch.ini";
        private NotifyIcon notifyIcon;
        private ContextMenuStrip sub_menu;
        private CustomMenuStrip main_menu;
        private ToolStripMenuItem menu_version;
        private ToolStripMenuItem menu_add_file;
        private ToolStripMenuItem menu_add_folder;
        private ToolStripMenuItem menu_add_path;
        private ToolStripMenuItem menu_delete;
        private ToolStripMenuItem menu_move_up;
        private ToolStripMenuItem menu_move_down;
        private ToolStripMenuItem menu_rename;
        private ToolStripMenuItem menu_close;
        private ToolStripMenuItem menu_exit;
        private ToolStripTextBox menu_path_edit;
        private System.Windows.Forms.Timer closeTimer;
        private bool suppressShow = false;

        private string explorer_path;
        private string browser_path;
        private string chrome_path = "";
        private List<ShortCutItem> shortCutList = new List<ShortCutItem>();
        private TextBox focus_text_box;

        /* コンストラクタ */
        public TTQLApplicationContext()
        {
            GetFtypeText();

            LoadShortCutIni();
            main_menu = new CustomMenuStrip(new Container());
            main_menu.Opened += main_menu_Opened; // Add event handler for opening
            main_menu.PreviewKeyDown += (s, e) =>
            {
                System.Diagnostics.Debug.WriteLine("main_nenu.keydown! : " + e.KeyCode);
                if (e.KeyCode == Keys.Escape)
                {
                    main_menu_closing();
                }
            };

            BuildMainMenuItems();

            ConstructSubMenuItems();
            BuildSubMenuItems();

            notifyIcon = new NotifyIcon()
            {
                Icon = Resources.TTQL_main,
                Visible = true
            };

            notifyIcon.MouseMove += IconOnMouseMove;
            notifyIcon.MouseClick += Icon_OnClick;
            main_menu.ContextMenuStrip = sub_menu;

            /* MouseMoveで表示したメインメニューを、10秒でCloseする */
            closeTimer = new System.Windows.Forms.Timer
            {
                Interval = Constants.MENU_HIDE_INTERVAL,
                Enabled = false
            };
            closeTimer.Tick += (s, e) =>
            {
                if (main_menu.Visible)
                {
                    main_menu_closing();
                }

                if (sub_menu.Visible)
                {
                    sub_menu.Hide();
                }
                LockMainMenuSelect(false);
                closeTimer.Stop();
                suppressShow = false;
            };
        }

        private void main_menu_Opened(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("main_menu_Opened");
            main_menu.BeginInvoke(new Action(() =>
            {
                if (focus_text_box != null)
                {
                    System.Diagnostics.Debug.WriteLine("main_menu_Opened focus on text_box!");
                    focus_text_box.Focus(); // Set focus to the text box if it exists
                    focus_text_box.SelectAll(); // Select all text in the text box
                }
            }));
        }

        /* メインメニューが開く際の前処理 */
        private void main_menu_Opening(object sender, CancelEventArgs e)
        {
        }

        /* サブメニュークローズ時の処理 */
        private void sub_menu_Closing(object sender, CancelEventArgs e)
        {
            LockMainMenuSelect(false);
        }

        /* サブメニューが開く際の前処理 */
        private void sub_menu_Opening(object sender, CancelEventArgs e)
        {
            ToolStripItem toolStripItem = selected_item();
            if (toolStripItem != null)
            {
                // If an item is selected, enable the delete option
                menu_delete.Enabled = true;
                menu_move_up.Enabled = true;
                menu_move_down.Enabled = true;
                menu_rename.Enabled = true;
            }
            else
            {
                // No item selected, disable the delete option
//              sub_menu.Items.Remove(menu_delete);
//              sub_menu.Items.Remove(menu_move_down);
//              sub_menu.Items.Remove(menu_move_up);
//              sub_menu.Items.Remove(menu_rename);
                menu_delete.Enabled = false;
                menu_move_down.Enabled = false;
                menu_move_up.Enabled = false;
                menu_rename.Enabled = false;
            }
            LockMainMenuSelect(true);
        }

        /* 環境変数からのパス取得 */
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

        /* サブメニューの初期設定 */
        private void ConstructSubMenuItems()
        {
            // Create the sub-menu items
            sub_menu = new ContextMenuStrip(new Container());
            sub_menu.Opening += sub_menu_Opening;
            sub_menu.Closing += sub_menu_Closing;
            var infoVersion = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            menu_version = new ToolStripMenuItem($"{Application.ProductName} v{infoVersion}")
            {
                Enabled = false
            };
            menu_add_file = new ToolStripMenuItem("Add File", null, AddFileShortcut);
            menu_add_folder = new ToolStripMenuItem("Add Folder", null, AddFolderShortcut);
            menu_add_path = new ToolStripMenuItem("Add Path", null, AddPathShortcut);
            menu_delete = new ToolStripMenuItem("Delete", null, DelShortcut);
            menu_exit = new ToolStripMenuItem("Exit", null, Exit);
            menu_move_down = new ToolStripMenuItem("Move Down", null, SubMenuMoveDown);
            menu_move_up = new ToolStripMenuItem("Move Up", null, SubMenuMoveUp);
            menu_rename = new ToolStripMenuItem("Rename", null, SubMenuRename);
            menu_close = new ToolStripMenuItem("Close", null, (s, e) =>
            {
                main_menu_closing();
            });
        }

        /* サブメニューアイテムの構築 */
        private void BuildSubMenuItems()
        {
            // Clear existing items
            sub_menu.Items.Clear();
            // Add shortcuts from the list
            sub_menu.Items.AddRange(new ToolStripItem[]
            {
                menu_version,
                new ToolStripSeparator()
            });

            sub_menu.Items.Add(menu_add_file);
            sub_menu.Items.Add(menu_add_folder);
            sub_menu.Items.Add(menu_add_path);
            sub_menu.Items.Add(menu_move_up);
            sub_menu.Items.Add(menu_move_down);
            sub_menu.Items.Add(menu_rename);
            sub_menu.Items.Add(menu_delete);
            sub_menu.Items.Add(menu_close);
            sub_menu.Items.Add(menu_exit);
        }



        /* パス指定のアイテム追加 */
        public void item_add_path(string name, string path)
        {
            AddShortCutItem(name, path, "");
            DelShortCutItem("");                              /* 編集中のショートカットを削除する */
            SaveShortCutIni();
            BuildMainMenuItems();
        }

        /* メインメニューアイテムのクリックハンドラ（カスタムアイテムからのコールバックでメニューを閉じる） */
        public void  item_on_click(object sender, MouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("item_on_click : " + e.Button);
            if (e.Button == MouseButtons.Left)
            {
                main_menu.BeginInvoke(new Action(() =>
                {
                    main_menu_closing();
                }));
            }
            else if (e.Button == MouseButtons.Right)
            {

            }
        }

        /* 選択されているショートカットアイテムでロック */
        private void LockMainMenuSelect(bool locked)
        {
            if (main_menu != null && main_menu.Items.Count > 0)
            {
                foreach (ToolStripItem item in main_menu.Items)
                {
                    if (item is ToolStripControlHost controlHost && controlHost.Control is CustomToolStripMenuItem customItem)
                    {
                        customItem.SetSelectLock(locked);
                    }
                }
            }
        }

        /* ToolStripItemからカスタムアイテムを取得する */
        private CustomToolStripMenuItem get_custom_item(ToolStripItem item)
        {
            if (item is ToolStripControlHost controlHost && controlHost.Control is CustomToolStripMenuItem customItem)
            {
                return customItem;
            }

            return null;
        }

        /* 選択されているショートカットアイテムを取得 */
        private ToolStripItem selected_item()
        {
            // Get the currently selected item in the main menu
            if (main_menu != null && main_menu.Items.Count > 0)
            {
                foreach (ToolStripItem item in main_menu.Items)
                {
                    if (item is ToolStripControlHost controlHost && controlHost.Control is CustomToolStripMenuItem customItem && customItem.IsSelected())
                    {
                        return item;
                    }
                }
            }
            return null; // No item is selected
        }

        /* メインメニューアイテムの構築 */
        private void BuildMainMenuItems()
        {
            // Clear existing items
            main_menu.Items.Clear();
            focus_text_box = null;
            // Add shortcuts from the list
#if VALID_CODE
            int index = 0;
            foreach (var item in shortCutList)
            {
                CustomToolStripMenuItem custom_item = new CustomToolStripMenuItem(item.Name, item.Icon?.ToBitmap(), item.Path, this)
                {
                    Tag = index++,
                    AutoSize = true,
                };
                ToolStripControlHost menuItem = new ToolStripControlHost(custom_item)
                {
                    AutoSize = true,
                    Padding =  Padding.Empty,
                    Margin =  Padding.Empty,
                };

                if (item.Path == "")
                {
                    focus_text_box = custom_item.Controls.OfType<FlowLayoutPanel>().FirstOrDefault().Controls.OfType<TextBox>().FirstOrDefault();
                }
                main_menu.Items.Add(menuItem);
            }
#else
            main_menu.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripSeparator(),
                new ToolStripSeparator()
            });
#endif
        }

        /* ショートカットアイテムの削除 */
        private void DelShortCutItem(string path)
        {
            foreach (var tmp_item in shortCutList)
            {
                if (tmp_item.Path.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    shortCutList.Remove(tmp_item);
                    return;
                }
            }
        }

        /* ショートカットアイテムの追加 */
        private void AddShortCutItem(string name, string path, string group)
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
                Group = group,
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

        /* ショートカット情報の読み込み */
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
                        if (parts[0].Trim() != "")
                        {
                            AddShortCutItem(parts[0].Trim(), parts[1].Trim(), "");
                        }
                    }
                    else if (parts.Length == 3)
                    {
                        if (parts[0].Trim() != "")
                        {
                            AddShortCutItem(parts[0].Trim(), parts[1].Trim(), parts[2].Trim());
                        }
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

        /* ショートカット情報の保存 */
        private void SaveShortCutIni()
        {
            try
            {
                StreamWriter sw = new StreamWriter(INI_FILE_NAME);
                sw.WriteLine("# TaskTrayQuickLaunch v" + Application.ProductVersion);
                foreach (var item in shortCutList)
                {
                    sw.WriteLine($"{item.Name}|{item.Path}|{item.Group}");
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

        /* メインメニュー表示 */
        private void DisplayMainMenu()
        {
            // Show the main menu at the current cursor position
            main_menu.Show(new Point(Cursor.Position.X - 220, Cursor.Position.Y - main_menu.Height - 32));
            main_menu.AutoClose = false;
            main_menu.Focus();
        }

        /* サブメニュー表示 */
        private void DisplaySubMenu()
        {
            // Show the main menu at the current cursor position
            sub_menu.Show(new Point(Cursor.Position.X, Cursor.Position.Y - sub_menu.Height - 32));
        }

        /* メインメニューのクローズ処理 */
        private void main_menu_closing()
        {
            System.Diagnostics.Debug.WriteLine("main_menu_closing()");
            if (focus_text_box != null)
            {
                DelShortCutItem("");                              /* 編集中のショートカットが残っている場合は、削除する */
                BuildMainMenuItems();
                main_menu.Close();
            }
            else
            {
                main_menu.AutoClose = true;
                LockMainMenuSelect(false);
                main_menu.Hide();
            }

            /* クローズして1秒間はMouseMoveによる再表示を抑止する */
            suppressShow = true;
            closeTimer.Stop();
            closeTimer.Interval = Constants.MENU_SUPPRESS_INTERVAL;
            closeTimer.Start();
        }

        /* NotifyIconのクリックハンドラ */
        private void Icon_OnClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                /* メインメニューが表示していたらクローズ、表示していなければ表示する */
                if (main_menu.Visible)
                {
                    main_menu_closing();
                }
                else
                {
                    // Show the main menu at the current cursor position
                    DisplayMainMenu();
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                DisplaySubMenu();
            }
        }

        /* NotifyIconのマウスムーブイベント */
        private void IconOnMouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (suppressShow)
            {
                return; // Do not show the main menu if suppressed
            }

            /* メインメニューもサブメニューも表示されていない場合、NotifyIcon上をマウスオーバーしただけでメインメニューを表示する。（10秒で閉じる）  */
            if ((main_menu.Visible == false) && (sub_menu.Visible == false))
            {
                DisplayMainMenu();
            }

            on_some_operation();
        }

        /* パス or URL指定のショートカット追加 */
        private void AddPathShortcut(object sender, EventArgs e)
        {
            AddShortCutItem("", "", "");
            BuildMainMenuItems();
            main_menu_Opened((object)sender, EventArgs.Empty);
            closeTimer.Stop();
        }

        /* フォルダへのショートカット追加 */
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
                        AddShortCutItem(paths[paths.Length -1], selectedPath, "");
                        SaveShortCutIni();
                        BuildMainMenuItems();
                    }
                }
            }
        }


        /* ファイルへのショートカット追加 */
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
                    AddShortCutItem(Path.GetFileNameWithoutExtension(filePath), filePath, "");
                    SaveShortCutIni();
                    BuildMainMenuItems();
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("No file selected.");
            }
        }

        /* 選択したアイテムのショートカットを削除 */
        private void DelShortcut(object sender, EventArgs e)
        {
            var item = selected_item();

            if (item != null)
            {
                var customItem = get_custom_item(item);
                // Get the path of the selected item
                string path = customItem.GetStartPath();
                if (!string.IsNullOrEmpty(path))
                {
                    // Remove the item from the list
                    DelShortCutItem(path);
                    SaveShortCutIni();
                    BuildMainMenuItems();
                }
            }
        }

        private int GetShortCutIndex(string path)
        {
            // Get the index of the shortcut item in the list
            for (int i = 0; i < shortCutList.Count; i++)
            {
                if (shortCutList[i].Path.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1; // Not found
        }

        private void SubMenuRename(object sender, EventArgs e)
        {

        }

        /* アイテムを上に移動 */
        private void SubMenuMoveUp(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("SubMenuMoveDown");
            var item = selected_item();
            if (item != null)
            {
                var custom_item = get_custom_item(item);
                int index = GetShortCutIndex(custom_item.GetStartPath());

                if (index > 0)
                {
                    // Move the item up in the list
                    var tmp_item = shortCutList[index];
                    shortCutList[index] = shortCutList[index - 1];
                    shortCutList[index - 1] = tmp_item;
                    // Rebuild the main menu items
                    SaveShortCutIni();
                    BuildMainMenuItems();
                }
            }
        }

        /* アイテムを下に移動 */
        private void SubMenuMoveDown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("SubMenuMoveDown");
            var item = selected_item();
            if (item != null)
            {
                var custom_item = get_custom_item(item);
                int index = GetShortCutIndex(custom_item.GetStartPath());

                if (index >= 0 && index < shortCutList.Count - 1)
                {
                    // Move the item down in the list
                    var tmp_item = shortCutList[index];
                    shortCutList[index] = shortCutList[index + 1];
                    shortCutList[index + 1] = tmp_item;
                    // Rebuild the main menu items
                    SaveShortCutIni();
                    BuildMainMenuItems();
                }
            }
        }

        public void on_focus(int index)
        {
            if (main_menu != null && main_menu.Items.Count > 0)
            {
                foreach (ToolStripItem manu_item in main_menu.Items)
                {
                    if (manu_item is ToolStripControlHost controlHost && controlHost.Control is CustomToolStripMenuItem customItem)
                    {
                        if (customItem.Tag is int selected && (selected != index))
                        {
                            customItem.SetSelected(false);
                        }
                    }
                }
            }

            on_some_operation();
        }

        private void on_some_operation()
        {
            if (focus_text_box == null)
            {
                /* マウスムーブで表示した場合は、10秒後に非表示にする */
                closeTimer.Stop();
                closeTimer.Interval = Constants.MENU_HIDE_INTERVAL;
                closeTimer.Start();
            }
            else
            {
                /* 入力モードの場合は、10秒タイムアウトしない */
                closeTimer.Stop();
            }
        }

        /* サブメニューのExit選択時のハンドラ */
        private void Exit(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }
    }

}