using Excel_Macros_UI.Model;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel;
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

namespace Excel_Macros_UI.View
{
    /// <summary>
    /// Interaction logic for ConsoleView.xaml
    /// </summary>
    public partial class ConsoleView : UserControl
    {
        public ConsoleView()
        {
            InitializeComponent();

            Routing.EventManager.GetInstance().ClearAllIOEvent += () =>
            {
                txtOutput.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => txtOutput.Clear()));
            };

            ConsoleModel.GetInstance().Output = new TextBoxWriter(txtOutput);
            ConsoleModel.GetInstance().Error = new TextBoxWriter(txtOutput);

            Routing.EventManager.ChangeIO(ConsoleModel.GetInstance().Output, ConsoleModel.GetInstance().Error);
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            txtOutput.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => txtOutput.Clear()));
        }
    }
}
