using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Excel_Macros_UI.Model
{
    public class ExplorerModel : ToolModel
    {
        private static ExplorerModel s_Instance;

        public static ExplorerModel GetInstance()
        {
            return s_Instance != null ? s_Instance : new ExplorerModel();
        }

        public ExplorerModel()
        {
            s_Instance = this;

            PreferredLocation = PaneLocation.Right;
            SelectedItem = null;
            ItemSource = new ObservableCollection<CustomTreeViewItem>();
        }

        #region SelectedItem

        private CustomTreeViewItem m_SelectedItem;
        public CustomTreeViewItem SelectedItem
        {
            get
            {
                return m_SelectedItem;
            }

            set
            {
                if (m_SelectedItem != value)
                {
                    m_SelectedItem = value;
                    OnPropertyChanged(nameof(SelectedItem));
                }
            }
        }

        #endregion

        #region ItemSource

        private ObservableCollection<CustomTreeViewItem> m_ItemSource;
        public ObservableCollection<CustomTreeViewItem> ItemSource
        {
            get
            {
                return m_ItemSource;
            }

            set
            {
                if (m_ItemSource != value)
                {
                    m_ItemSource = value;
                    OnPropertyChanged(nameof(ItemSource));
                }
            }
        }

        #endregion
    }
}
