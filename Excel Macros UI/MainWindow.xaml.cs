/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.3
 * The main window, hosting all the UI
 */

using Excel_Macros_INTEROP;
using Excel_Macros_UI.ViewModel;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml;
using Xceed.Wpf.AvalonDock.Layout.Serialization;

namespace Excel_Macros_UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private static MainWindow s_Instance;

        public MainWindow()
        {
            InitializeComponent();
            s_Instance = this;

            this.Loaded += MainWindow_Loaded;
        }

        public static MainWindow GetInstance()
        {
            return s_Instance;
        }
        
        public static void CreateInstance()
        {
            MainWindow mw = new MainWindow();
            PrimaryViewModel pvm = new PrimaryViewModel();
            mw.DataContext = pvm;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        { 
            string layout = Properties.Settings.Default.AvalonLayout;

            if (String.IsNullOrEmpty(layout))
                return;

            XmlLayoutSerializer serializer = new XmlLayoutSerializer(DockingManager_DockManager);

            StringReader stringReader = new StringReader(layout);
            XmlReader xmlReader = XmlReader.Create(stringReader);

            serializer.Deserialize(xmlReader);

            xmlReader.Close();
            stringReader.Close();
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            XmlLayoutSerializer serializer = new XmlLayoutSerializer(DockingManager_DockManager);
            StringWriter stringWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(stringWriter);

            serializer.Serialize(xmlWriter);

            xmlWriter.Flush();
            stringWriter.Flush();

            string layout = stringWriter.ToString();

            xmlWriter.Close();
            stringWriter.Close();

            Properties.Settings.Default.AvalonLayout = layout;
            Properties.Settings.Default.Save();

            this.Hide();
        }

        private void MetroWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {

        }

        #region EventManager Event Function Callbacks

        public void ShowWindow()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                Main.GetInteropDispatcher().Invoke(() => Main.SetExcelInteractive(false));
                Show();
                Focus();
                Activate();
            }));
        }

        public void HideWindow()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                Main.GetInteropDispatcher().Invoke(() => Main.SetExcelInteractive(true));
                Properties.Settings.Default.Save();
                Close();
            }));
        }

        public void TryFocus()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { Focus(); }));
        }

        public void DisplayOkMessage(string message, string caption)
        {
            if (IsVisible)
                Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(async () => { await this.ShowMessageAsync(caption, message, MessageDialogStyle.Affirmative); }));
            else
                System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK);
        }

        public void DisplayYesNoMessage(string message, string caption, Action<bool> OnReturn)
        {
            if (IsVisible)
            {
                Dispatcher.Invoke(async () =>
                {
                    bool result = (await this.ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegative)) == MessageDialogResult.Affirmative;
                    OnReturn?.Invoke(result);
                });
            }
            else
                OnReturn?.Invoke(System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK) == System.Windows.Forms.DialogResult.OK);
        }

        #endregion
    }
}
