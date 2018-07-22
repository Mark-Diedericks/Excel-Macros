/*
 * Mark Diedericks
 * 19/07/2018
 * Version 1.0.4
 * Settings menu basic view logic
 */
using System;
using System.Collections.Generic;
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
using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.ViewModel;
using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

namespace Excel_Macros_UI.View
{
    /// <summary>
    /// Interaction logic for SettingsMenuView.xaml
    /// </summary>
    public partial class SettingsMenuView : Flyout
    {
        public SettingsMenuView()
        {
            InitializeComponent();
            SetActiveSettingsPage(SettingsPage.Editor);

            bool light = Properties.Settings.Default.Theme == "Light";
            rdbtnThemeLight.IsChecked = light;
            rdbtnThemeDark.IsChecked = !light;
        }

        public ResourceDictionary ThemeDictionary
        {
            get
            {
                return Resources.MergedDictionaries[1];
            }
        }

        private enum SettingsPage
        {
            Editor = 0,
            Libraries = 1,
            Ribbon = 2
        }

        public void CreateSettingsMenu()
        {
            //CREATE EDITOR SETTINGS TAB

            //Nothing Needs to be done

            //CREATE LIBRARY SETTINGS TAB



            //CREATE RIBBON SETTINGS TAB

            Dictionary<Guid, IMacro> macros = Main.GetMacros();
            listSettingsRibbonMacro.Items.Clear();
            foreach (Guid id in macros.Keys)
                listSettingsRibbonMacro.Items.Add(CreateRibbonMacroListItem(id, macros[id]));

            //Set default page
            SetActiveSettingsPage(SettingsPage.Editor);
        }

        private ListViewItem CreateRibbonMacroListItem(Guid id, IMacro macro)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            ListViewItem lvi = new ListViewItem();
            CheckBox cb = new CheckBox();

            cb.Content = md.name;
            cb.ToolTip = md.relativepath;
            cb.IsChecked = Main.IsRibbonMacro(id);

            cb.Checked += delegate (object sender, RoutedEventArgs args) { Main.AddRibbonMacro(id); };
            cb.Unchecked += delegate (object sender, RoutedEventArgs args) { Main.RemoveRibbonMacro(id); };

            lvi.Content = cb;

            return lvi;
        }

        private void btnSettingsEditor_Click(object sender, RoutedEventArgs e)
        {
            SetActiveSettingsPage(SettingsPage.Editor);
        }

        private void btnSettingsLibraries_Click(object sender, RoutedEventArgs e)
        {
            SetActiveSettingsPage(SettingsPage.Libraries);
        }

        private void btnSettingsRibbon_Click(object sender, RoutedEventArgs e)
        { 

            Dictionary<Guid, IMacro> macros = Main.GetMacros();
            lblEmpty.Visibility = macros.Count() > 0 ? Visibility.Hidden : Visibility.Visible;

            listSettingsRibbonMacro.Items.Clear();
            foreach (Guid id in macros.Keys)
                listSettingsRibbonMacro.Items.Add(CreateRibbonMacroListItem(id, macros[id]));

            SetActiveSettingsPage(SettingsPage.Ribbon);
        }

        private void SetActiveSettingsPage(SettingsPage page)
        {
            if (page == SettingsPage.Editor)
            {
                gridSettingsEditor.Visibility = Visibility.Visible;
                gridSettingsLibraries.Visibility = Visibility.Hidden;
                gridSettingsRibbon.Visibility = Visibility.Hidden;

                btnSettingsEditor.IsChecked = true;
                btnSettingsLibraries.IsChecked = false;
                btnSettingsRibbon.IsChecked = false;
            }

            if (page == SettingsPage.Libraries)
            {
                gridSettingsEditor.Visibility = Visibility.Hidden;
                gridSettingsLibraries.Visibility = Visibility.Visible;
                gridSettingsRibbon.Visibility = Visibility.Hidden;

                btnSettingsEditor.IsChecked = false;
                btnSettingsLibraries.IsChecked = true;
                btnSettingsRibbon.IsChecked = false;
            }

            if (page == SettingsPage.Ribbon)
            {
                gridSettingsEditor.Visibility = Visibility.Hidden;
                gridSettingsLibraries.Visibility = Visibility.Hidden;
                gridSettingsRibbon.Visibility = Visibility.Visible;

                btnSettingsEditor.IsChecked = false;
                btnSettingsLibraries.IsChecked = false;
                btnSettingsRibbon.IsChecked = true;
            }
        }

        private void rdbtnThemeLight_Checked(object sender, RoutedEventArgs e)
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().SetTheme("Light");
        }

        private void rdbtnThemeDark_Checked(object sender, RoutedEventArgs e)
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().SetTheme("Dark");
        }
    }
}
