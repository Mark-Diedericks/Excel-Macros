/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Primary entry point into the application -> auto-generated
 */

using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Excel_Macros_UI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            MainWindow mw = new MainWindow();
            mw.DataContext = new PrimaryViewModel();
            MainWindow = mw;

        }
    }
}
