using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Excel_Macros_UI.Model
{
    /// <summary>
    /// Custom data structure for holding the temporary data of each tree view item
    /// </summary>
    public class DataTreeViewItem
    {
        public int level;
        public bool folder;
        public string name;
        public Guid macro;
        public string root;
        public List<DataTreeViewItem> children;
    }

    /// <summary>
    /// Custom Data Strcture for Displaying Tree View Items
    /// </summary>
    public class DisplayableTreeViewItem : ViewModel.Base.ViewModel
    {
        public DisplayableTreeViewItem()
        {
            Header = "";
            IsExpanded = false;
            IsFolder = false;
            IsInputting = false;
            Root = "";
            ID = Guid.Empty;
            Parent = null;
            Items = new ObservableCollection<DisplayableTreeViewItem>();
        }

        public void Selected(object sender, RoutedEventArgs args)
        {
            SelectedEvent?.Invoke(sender, args);
        }

        public void DoubleClick(object sender, MouseButtonEventArgs args)
        {
            DoubleClickEvent?.Invoke(sender, args);
        }

        public void RightClick(object sender, MouseButtonEventArgs args)
        {
            RightClickEvent?.Invoke(sender, args);
        }

        public void FocusLost(object sender, RoutedEventArgs args)
        {
            FocusLostEvent?.Invoke(sender, args);
        }

        public void KeyUp(object sender, KeyEventArgs args)
        {
            KeyUpEvent?.Invoke(sender, args);
        }
        
        #region IsInputting & IsDisplaying
        private bool m_IsInputting;
        public bool IsInputting
        {
            get
            {
                return m_IsInputting;
            }
            set
            {
                if (m_IsInputting != value)
                {
                    m_IsInputting = value;

                    OnPropertyChanged(nameof(IsDisplaying));
                    OnPropertyChanged(nameof(IsInputting));
                }
            }
        }
        public bool IsDisplaying
        {
            get
            {
                return !m_IsInputting;
            }
            set
            {
                if ((!m_IsInputting) != value)
                {
                    m_IsInputting = !value;

                    OnPropertyChanged(nameof(IsDisplaying));
                    OnPropertyChanged(nameof(IsInputting));
                }
            }
        }
        #endregion

        #region Header
        private string m_Header;
        public string Header
        {
            get
            {
                return m_Header;
            }
            set
            {
                if(m_Header != value)
                {
                    m_Header = value;
                    OnPropertyChanged(nameof(Header));
                }
            }
        }
        #endregion
        #region IsExpanded
        private bool m_IsExpaned;
        public bool IsExpanded
        {
            get
            {
                return m_IsExpaned;
            }
            set
            {
                if (m_IsExpaned != value)
                {
                    m_IsExpaned = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }
        }
        #endregion 
        #region IsFolder
        private bool m_IsFolder;
        public bool IsFolder
        {
            get
            {
                return m_IsFolder;
            }
            set
            {
                if (m_IsFolder != value)
                {
                    m_IsFolder = value;
                    OnPropertyChanged(nameof(IsFolder));
                }
            }
        }
        #endregion 
        #region Root
        private string m_Root;
        public string Root
        {
            get
            {
                return m_Root;
            }
            set
            {
                if (m_Root != value)
                {
                    m_Root = value;
                    OnPropertyChanged(nameof(Root));
                }
            }
        }
        #endregion 
        #region ID
        private Guid m_ID;
        public Guid ID
        {
            get
            {
                return m_ID;
            }
            set
            {
                if (m_ID != value)
                {
                    m_ID = value;
                    OnPropertyChanged(nameof(ID));
                }
            }
        }
        #endregion 
        #region Parent
        private DisplayableTreeViewItem m_Parent;
        public DisplayableTreeViewItem Parent
        {
            get
            {
                return m_Parent;
            }
            set
            {
                if (m_Parent != value)
                {
                    m_Parent = value;
                    OnPropertyChanged(nameof(Parent));
                }
            }
        }
        #endregion 
        #region Items
        private ObservableCollection<DisplayableTreeViewItem> m_Items;
        public ObservableCollection<DisplayableTreeViewItem> Items
        {
            get
            {
                return m_Items;
            }
            set
            {
                if (m_Items != value)
                {
                    m_Items = value;
                    OnPropertyChanged(nameof(Items));
                }
            }
        }
        #endregion

        #region SelectedEvent
        private Action<object, RoutedEventArgs> m_SelectedEvent;
        public Action<object, RoutedEventArgs> SelectedEvent
        {
            get
            {
                return m_SelectedEvent;
            }
            set
            {
                if (m_SelectedEvent != value)
                {
                    m_SelectedEvent = value;
                    OnPropertyChanged(nameof(SelectedEvent));
                }
            }
        }
        #endregion 
        #region DoubleClickEvent
        private Action<object, MouseButtonEventArgs> m_DoubleClickEvent;
        public Action<object, MouseButtonEventArgs> DoubleClickEvent
        {
            get
            {
                return m_DoubleClickEvent;
            }
            set
            {
                if (m_DoubleClickEvent != value)
                {
                    m_DoubleClickEvent = value;
                    OnPropertyChanged(nameof(DoubleClickEvent));
                }
            }
        }
        #endregion 
        #region RightClickEvent
        private Action<object, MouseButtonEventArgs> m_RightClickEvent;
        public Action<object, MouseButtonEventArgs> RightClickEvent
        {
            get
            {
                return m_RightClickEvent;
            }
            set
            {
                if (m_RightClickEvent != value)
                {
                    m_RightClickEvent = value;
                    OnPropertyChanged(nameof(RightClickEvent));
                }
            }
        }
        #endregion 

        #region FocusLostEvent
        private Action<object, RoutedEventArgs> m_FocusLostEvent;
        public Action<object, RoutedEventArgs> FocusLostEvent
        {
            get
            {
                return m_FocusLostEvent;
            }
            set
            {
                if (m_FocusLostEvent != value)
                {
                    m_FocusLostEvent = value;
                    OnPropertyChanged(nameof(FocusLostEvent));
                }
            }
        }
        #endregion 
        #region KeyUpEvent
        private Action<object, KeyEventArgs> m_KeyUpEvent;
        public Action<object, KeyEventArgs> KeyUpEvent
        {
            get
            {
                return m_KeyUpEvent;
            }
            set
            {
                if (m_KeyUpEvent != value)
                {
                    m_KeyUpEvent = value;
                    OnPropertyChanged(nameof(KeyUpEvent));
                }
            }
        }
        #endregion 
    }

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
