/*
 * Mark Diedericks
 * 26/06/2018
 * Version 1.0.1
 * File explorer view model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.View;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel
{
    public class ExplorerViewModel : ToolViewModel
    {
        public ExplorerViewModel()
        {
            Model = new ExplorerModel();
        }

        #region Model

        public new ExplorerModel Model
        {
            get
            {
                return (ExplorerModel)base.Model;
            }

            set
            {
                if (((ExplorerModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region SelectedItem

        public TreeViewItem SelectedItem
        {
            get
            {
                return Model.SelectedItem;
            }
            set
            {
                if (Model.SelectedItem != value)
                {
                    Model.SelectedItem = value;
                    OnPropertyChanged(nameof(SelectedItem));
                }
            }
        }

        #endregion

        #region ItemSource

        public ItemCollection ItemSource
        {
            get
            {
                return Model.ItemSource;
            }
            set
            {
                if (Model.ItemSource != value)
                {
                    Model.ItemSource = value;
                    OnPropertyChanged(nameof(ItemSource));
                }
            }
        }

        #endregion
        
    }
}
