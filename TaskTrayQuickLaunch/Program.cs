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
        private NotifyIcon notifyIcon;
        private ContextMenuStrip sub_menu;
        private ContextMenuStrip main_menu;
        private bool main_menu_show;
        private string current_dir;

        public TTQLApplicationContext()
        {
            Application.ApplicationExit += new EventHandler(OnApplicationExit);
            current_dir = Directory.GetCurrentDirectory();


            LoadShortCutIni();
            main_menu = new ContextMenuStrip(new Container());
            main_menu.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripSeparator(),
                new ToolStripSeparator(),
            });

            sub_menu = new ContextMenuStrip(new Container());
            sub_menu.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripMenuItem($"{Application.ProductName} v{Application.ProductVersion}")
                {
                    Enabled = false
                },
                new ToolStripSeparator(),
                new ToolStripMenuItem("í«â¡(add)", null, Exit),
                new ToolStripMenuItem("çÌèú(delete)", null, Exit),
                new ToolStripMenuItem("Exit", null, Exit)
            });

            notifyIcon = new NotifyIcon()
            {
                Icon = Resources.TTQL_main,
                Visible = true
            };

            main_menu_show = false;
            notifyIcon.MouseMove += OnMouseMove;
            notifyIcon.ContextMenuStrip = sub_menu;
        }

        private void LoadShortCutIni()
        {

        }

        private void OnMouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
//          notifyIcon.Text = Directory.GetCurrentDirectory();
            if (main_menu_show == false)
            {
                main_menu.Show(Cursor.Position);
                main_menu_show = true;
                
            }

        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

        }

        private void Exit(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }
    }

}