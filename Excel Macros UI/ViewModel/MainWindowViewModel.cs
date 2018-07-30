﻿/*
 * Mark Diedericks
 * 24/07/2018
 * Version 1.0.4
 * Primary view model for handling main window's views
 */

using Excel_Macros_UI.Model;
using Excel_Macros_UI.Themes;
using Excel_Macros_UI.ViewModel.Base;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Excel_Macros_UI.ViewModel
{
    public class MainWindowViewModel : Base.ViewModel, IThemeManager
    {
        private static MainWindowViewModel s_Instance;

        public MainWindowViewModel()
        {
            s_Instance = this;

            Model = new MainWindowModel();
            DockManager = new DockManagerViewModel(Properties.Settings.Default.OpenDocuments);
        }

        public MainWindowViewModel GetInstance()
        {
            return s_Instance;
        }

        #region Model

        private MainWindowModel m_Model;
        public MainWindowModel Model
        {
            get
            {
                return m_Model;
            }
            set
            {
                if(m_Model != value)
                {
                    m_Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region DockManager

        public DockManagerViewModel DockManager
        {
            get
            {
                return Model.DockManager;
            }
            set
            {
                if(Model.DockManager != value)
                {
                    Model.DockManager = value;
                    OnPropertyChanged(nameof(DockManager));
                }
            }
        }

        #endregion

        #region IsClosing

        public bool IsClosing
        {
            get
            {
                return Model.IsClosing;
            }
            set
            {
                if(Model.IsClosing != value)
                {
                    Model.IsClosing = value;
                    OnPropertyChanged(nameof(IsClosing));
                }
            }
        }

        #endregion

        #region SettingsMenu

        public SettingsMenuViewModel SettingsMenu
        {
            get
            {
                return Model.SettingsMenu;
            }
            set
            {
                if(Model.SettingsMenu != value)
                {
                    Model.SettingsMenu = value;
                    OnPropertyChanged(nameof(SettingsMenu));
                }
            }
        }

        #endregion

        #region Themes

        public ObservableCollection<ITheme> Themes
        {
            get
            {
                return Model.Themes;
            }
            set
            {
                if(Model.Themes != value)
                {
                    Model.Themes = value;
                    OnPropertyChanged(nameof(Themes));
                }
            }
        }

        #endregion

        #region ActiveTheme

        public ITheme ActiveTheme
        {
            get
            {
                return Model.ActiveTheme;
            }
            set
            {
                if(Model.ActiveTheme != value)
                {
                    Model.ActiveTheme = value;
                    OnPropertyChanged(nameof(ActiveTheme));
                }
            }
        }

        #endregion

        #region DocumentContextMenu

        public ContextMenu DocumentContextMenu
        {
            get
            {
                return Model.DocumentContextMenu;
            }
            set
            {
                if(Model.DocumentContextMenu != value)
                {
                    Model.DocumentContextMenu = value;
                    OnPropertyChanged(nameof(DocumentContextMenu));
                }
            }
        }

        #endregion

        #region AnchorableContextMenu

        public ContextMenu AnchorableContextMenu
        {
            get
            {
                return Model.AnchorableContextMenu;
            }
            set
            {
                if (Model.AnchorableContextMenu != value)
                {
                    Model.AnchorableContextMenu = value;
                    OnPropertyChanged(nameof(AnchorableContextMenu));
                }
            }
        }

        #endregion

        #region ActiveDocument

        public DocumentViewModel ActiveDocument
        {
            get
            {
                return Model.ActiveDocument;
            }
            set
            {
                if (Model.ActiveDocument != value)
                {
                    Model.ActiveDocument = value;
                    OnPropertyChanged(nameof(ActiveDocument));
                }
            }
        }

        #endregion

        #region SelectedExecutionIndex

        public int SelectedExecutionIndex
        {
            get
            {
                return Model.SelectedExecutionIndex;
            }
            set
            {
                if(Model.SelectedExecutionIndex != value)
                {
                    Model.SelectedExecutionIndex = value;
                    OnPropertyChanged(nameof(SelectedExecutionIndex));
                    OnPropertyChanged(nameof(AsyncExecution));
                }
            }
        }

        #endregion

        #region AsyncExecution

        public bool AsyncExecution
        {
            get
            {
                return Model.AsyncExecution;
            }
            set
            {
                if(Model.AsyncExecution != value)
                {
                    Model.AsyncExecution = value;
                    OnPropertyChanged(nameof(AsyncExecution));
                }
            }
        }

        #endregion
    }
}
