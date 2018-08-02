/*
 * Mark Diedericks
 * 30/06/2018
 * Version 1.0.3
 * File explorer model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
            
            SelectedItem = null;
            ItemSource = new ObservableCollection<DisplayableTreeViewItem>();
            LabelVisibility = Visibility.Hidden;
        }

        #region SelectedItem

        private DisplayableTreeViewItem m_SelectedItem;
        public DisplayableTreeViewItem SelectedItem
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

        private ObservableCollection<DisplayableTreeViewItem> m_ItemSource;
        public ObservableCollection<DisplayableTreeViewItem> ItemSource
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

        #region LabelVisibility

        private Visibility m_LabelVisibility;
        public Visibility LabelVisibility
        {
            get
            {
                return m_LabelVisibility;
            }
            set
            {
                if(m_LabelVisibility != value)
                {
                    m_LabelVisibility = value;
                    OnPropertyChanged(nameof(LabelVisibility));
                }
            }
        }

        #endregion
    }
}
