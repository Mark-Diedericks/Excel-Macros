﻿using Excel_Macros_UI.Themes;
using Excel_Macros_UI.ViewModel;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Excel_Macros_UI.Model
{
    public class MainWindowModel : Base.Model
    {

        public MainWindowModel()
        {
            IsShown = false;
            IsFocused = false;
            IsClosing = false;
            IsExecuting = false;
            Themes = new ObservableCollection<ITheme>();
        }

        #region DockManager

        private DockManagerViewModel m_DockManager;
        public DockManagerViewModel DockManager
        {
            get
            {
                return m_DockManager;
            }
            set
            {
                if (m_DockManager != value)
                {
                    m_DockManager = value;
                    OnPropertyChanged(nameof(DockManager));
                }
            }
        }

        #endregion

        #region IsShown

        private bool m_IsShown;
        public bool IsShown
        {
            get
            {
                return m_IsShown;
            }
            set
            {
                if(m_IsShown != value)
                {
                    m_IsShown = value;
                    OnPropertyChanged(nameof(IsShown));
                }
            }
        }

        #endregion

        #region IsFocused

        private bool m_IsFocused;
        public bool IsFocused
        {
            get
            {
                return m_IsFocused;
            }
            set
            {
                if (m_IsFocused != value)
                {
                    m_IsFocused = value;
                    OnPropertyChanged(nameof(IsFocused));
                }
            }
        }

        #endregion

        #region IsClosing

        private bool m_IsClosing;
        public bool IsClosing
        {
            get
            {
                return m_IsClosing;
            }
            set
            {
                if (m_IsClosing != value)
                {
                    m_IsClosing = value;
                    OnPropertyChanged(nameof(IsClosing));
                }
            }
        }

        #endregion

        #region IsExecuting

        private bool m_IsExecuting;
        public bool IsExecuting
        {
            get
            {
                return m_IsExecuting;
            }
            set
            {
                if(m_IsExecuting != value)
                {
                    m_IsExecuting = value;
                    OnPropertyChanged(nameof(IsExecuting));
                }
            }
        }

        #endregion

        #region SettingsMenu

        private SettingsMenuViewModel m_SettingsMenu;
        public SettingsMenuViewModel SettingsMenu
        {
            get
            {
                return m_SettingsMenu;
            }
            set
            {
                if (m_SettingsMenu != value)
                {
                    m_SettingsMenu = value;
                    OnPropertyChanged(nameof(SettingsMenu));
                }
            }
        }

        #endregion

        #region Themes

        private ObservableCollection<ITheme> m_Themes;
        public ObservableCollection<ITheme> Themes
        {
            get
            {
                return m_Themes;
            }
            set
            {
                if (m_Themes != value)
                {
                    m_Themes = value;
                    OnPropertyChanged(nameof(Themes));
                }
            }
        }

        #endregion

        #region ActiveTheme

        private ITheme m_ActiveTheme;
        public ITheme ActiveTheme
        {
            get
            {
                return m_ActiveTheme;
            }
            set
            {
                if(m_ActiveTheme != value)
                {
                    m_ActiveTheme = value;
                    OnPropertyChanged(nameof(ActiveTheme));
                }
            }
        }

        #endregion

        #region DocumentContextMenu

        private ContextMenu m_DocumentContextMenu;
        public ContextMenu DocumentContextMenu
        {
            get
            {
                return m_DocumentContextMenu;
            }
            set
            {
                if(m_DocumentContextMenu != value)
                {
                    m_DocumentContextMenu = value;
                    OnPropertyChanged(nameof(DocumentContextMenu));
                }
            }
        }

        #endregion

        #region AnchorableContextMenu

        private ContextMenu m_AnchorableContextMenu;
        public ContextMenu AnchorableContextMenu
        {
            get
            {
                return m_AnchorableContextMenu;
            }
            set
            {
                if (m_AnchorableContextMenu != value)
                {
                    m_AnchorableContextMenu = value;
                    OnPropertyChanged(nameof(AnchorableContextMenu));
                }
            }
        }

        #endregion

        #region AsyncExecution

        public bool AsyncExecution
        {
            get
            {
                return Properties.Settings.Default.ExecutionTypeIndex == 0;
            }
            set
            {
                if((Properties.Settings.Default.ExecutionTypeIndex == 0) != value)
                {
                    Properties.Settings.Default.ExecutionTypeIndex = value ? 0 : 1;
                    OnPropertyChanged(nameof(AsyncExecution));
                }
            }
        }

        #endregion

    }
}