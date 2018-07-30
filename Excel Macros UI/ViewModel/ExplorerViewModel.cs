/*
 * Mark Diedericks
 * 30/06/2018
 * Version 1.0.3
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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace Excel_Macros_UI.ViewModel
{
    public class ExplorerViewModel : ToolViewModel
    {
        private bool m_IsCreating = false;

        public ExplorerViewModel()
        {
            Model = new ExplorerModel();

            if (Routing.EventManager.IsLoaded())
                Initialize();
            else
                Routing.EventManager.ApplicationLoaded += Initialize;
        }

        private void Focus()
        {
            FocusEvent?.Invoke();
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

        public DisplayableTreeViewItem SelectedItem
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

        public ObservableCollection<DisplayableTreeViewItem> ItemSource
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

        #region LabelVisibility

        public Visibility LabelVisibility
        {
            get
            {
                return Model.LabelVisibility;
            }
            set
            {
                if(Model.LabelVisibility != value)
                {
                    Model.LabelVisibility = value;
                    OnPropertyChanged(nameof(LabelVisibility));
                }
            }
        }

        #endregion

        #region FocusEvent
        private Action m_FocusEvent;
        public Action FocusEvent
        {
            get
            {
                return m_FocusEvent;
            }
            set
            {
                if (m_FocusEvent != value)
                {
                    m_FocusEvent = value;
                    OnPropertyChanged(nameof(FocusEvent));
                }
            }
        }
        #endregion

        #region Tree View Population Through Recursion & Sorting

        private void Rename(DisplayableTreeViewItem parent, DisplayableTreeViewItem item)
        {
            Remove(parent, item);
            Add(parent, item);

            UpdateRibbon();
        }

        private void Add(DisplayableTreeViewItem parent, DisplayableTreeViewItem item)
        {
            ObservableCollection<DisplayableTreeViewItem> siblings;

            if (parent == null)
                siblings = ItemSource;
            else
                siblings = parent.Items;
            
            int index = FindIndex(siblings, item);

            if (index == -1)
                index = 0;

            if (parent == null)
                ItemSource.Insert(index, item);
            else
                parent.Items.Insert(index, item);

            CheckVisibility();
        }

        private void Remove(DisplayableTreeViewItem parent, DisplayableTreeViewItem item)
        {
            if (parent == null)
                ItemSource.Remove(item);
            else
                parent.Items.Remove(item);

            CheckVisibility();
        }

        private void CheckVisibility()
        {
            if (ItemSource.Count == 0)
                LabelVisibility = Visibility.Visible;
            else
                LabelVisibility = Visibility.Hidden;

            UpdateRibbon();
        }

        private void UpdateRibbon()
        {
            SettingsMenuModel.SetRibbonItems(ItemSource);
        }

        private int FindIndex(ObservableCollection<DisplayableTreeViewItem> items, DisplayableTreeViewItem item)
        {
            int count = items.Count - 1;

            int low = 0;
            int high = count;
            int mid = (low + high) / 2;

            while(low <= high)
            {
                mid = (low + high) / 2;
                int pos = String.Compare(item.Header, items[mid].Header, true);
                
                if (pos < 0)
                    high = mid - 1;
                else if (pos > 0)
                    low = mid + 1;
                else
                    return mid;
            }

            if (String.Compare(item.Header, items[mid].Header, true) > 0)
                mid++;

            return mid;
        }

        private int Partition(ObservableCollection<DisplayableTreeViewItem> items, int start, int end)
        {
            int pivot = end;
            int i = start;
            int j = end;

            DisplayableTreeViewItem temp;
            while(i < j)
            {
                while (i < end && String.Compare(items[i].Header, items[pivot].Header, true) < 0)
                    i++;

                while (j > start && String.Compare(items[j].Header, items[pivot].Header, true) > 0)
                    j--;

                if(i < j)
                {
                    temp = items[i];
                    items[i] = items[j];
                    items[j] = temp;
                }
            }

            temp = items[pivot];
            items[pivot] = items[j];
            items[j] = temp;

            return j;
        }

        private void QuickSort(ref ObservableCollection<DisplayableTreeViewItem> items, int start, int end)
        {
            if(start < end)
            {
                int pivot = Partition(items, start, end);
                QuickSort(ref items, start, pivot - 1);
                QuickSort(ref items, pivot + 1, end);
            }
        }

        private void Sort()
        {
#if false
            List<DisplayableTreeViewItem> items = ItemSource.ToList<DisplayableTreeViewItem>();
            items.Sort();

            ItemSource.Clear();
            for (int i = 0; i < items.Count; i++)
                ItemSource.Insert(i, Sort(items[i]));
#else
            ObservableCollection<DisplayableTreeViewItem> items = ItemSource;
            QuickSort(ref items, 0, ItemSource.Count - 1);

            for (int i = 0; i < items.Count; i++)
                if (items[i].Items.Count > 0)
                    items[i] = Sort(items[i]);

            ItemSource = items;
#endif
        }

        private DisplayableTreeViewItem Sort(DisplayableTreeViewItem item)
        {
#if false
            List<DisplayableTreeViewItem> items = item.Items.ToList<DisplayableTreeViewItem>();
            items.Sort();

            item.Items.Clear();
            for (int i = 0; i < items.Count; i++)
                item.Items.Insert(i, Sort(items[i]));

            return item;
#else
            ObservableCollection<DisplayableTreeViewItem> items = item.Items;
            QuickSort(ref items, 0, item.Items.Count - 1);

            for (int i = 0; i < items.Count; i++)
                if (items[i].Items.Count > 0)
                    items[i] = Sort(items[i]);

            item.Items = items;
            return item;
#endif
        }

        private void Initialize()
        {
            Dictionary<Guid, IMacro>.KeyCollection keys = Main.GetMacros().Keys;
            HashSet<DataTreeViewItem> items = CreateTreeViewItemStructure(keys.ToList<Guid>());

            foreach (DataTreeViewItem item in items)
            {
                DisplayableTreeViewItem tvi = CreateTreeViewItem(null, item);

                if (tvi != null)
                    ItemSource.Add(tvi);
            }

            Sort();
            CheckVisibility();
            Main.GetInstance().OnMacroCountChanged += CheckVisibility;
        }

        private DisplayableTreeViewItem AddEvent(DisplayableTreeViewItem item)
        {
            if (item == null)
                return item;

            item.RightClickEvent += TreeViewItem_OnPreviewMouseRightButtonDown;

            for (int i = 0; i < item.Items.Count; i++)
                item.Items[i] = AddEvent(item.Items[i] as DisplayableTreeViewItem);

            return item;
        }

        private DisplayableTreeViewItem CreateTreeViewItem(DisplayableTreeViewItem parent, DataTreeViewItem data)
        {
            DisplayableTreeViewItem item = new DisplayableTreeViewItem();
            item.Header = data.name;
            item.IsFolder = data.folder;
            item.IsExpanded = false;
            item.Parent = parent;
            item.IsRibbonMacro = Main.IsRibbonMacro(data.macro);

            if (data.children != null)
            {
                foreach (DataTreeViewItem child in data.children)
                {
                    DisplayableTreeViewItem node = CreateTreeViewItem(item, child);

                    if (node != null)
                        item.Items.Add(node);
                }
            }

            if (!data.folder)
            {
                item.ID = data.macro;

                item.SelectedEvent += delegate (object sender, RoutedEventArgs args) { SelectedItem = item; };
                item.DoubleClickEvent += delegate (object sender, MouseButtonEventArgs args) { OpenMacro(data.macro); args.Handled = true; };
            }

            item.Root = data.root;
            item.RightClickEvent += TreeViewItem_OnPreviewMouseRightButtonDown;

            return item;
        }

        private HashSet<DataTreeViewItem> CreateTreeViewItemStructure(List<Guid> macros)
        {
            HashSet<DataTreeViewItem> items = new HashSet<DataTreeViewItem>();

            foreach (Guid id in macros)
            {
                string path = Regex.Replace(Main.GetDeclaration(id).relativepath, @"/+", System.IO.Path.DirectorySeparatorChar.ToString());
                string[] fileitems = path.Split(System.IO.Path.DirectorySeparatorChar).Where<string>(x => !String.IsNullOrEmpty(x)).ToArray<string>();

                if (fileitems.Any())
                {
                    DataTreeViewItem root = items.FirstOrDefault(x => x.name.Equals(fileitems[0]) && x.level.Equals(1));

                    if (root == null)
                    {
                        root = new DataTreeViewItem() { level = 1, name = fileitems[0], macro = id, root = "/", folder = !fileitems[0].ToLower().EndsWith(".ipy"), children = new List<DataTreeViewItem>() };
                        items.Add(root);
                    }

                    if (fileitems.Length > 1)
                    {
                        DataTreeViewItem parent = root;
                        int lev = 2;

                        for (int i = 1; i < fileitems.Length; i++)
                        {
                            DataTreeViewItem child = parent.children.FirstOrDefault(x => x.name.Equals(fileitems[i]) && x.level.Equals(lev));

                            if (child == null)
                            {
                                child = new DataTreeViewItem() { level = lev, name = fileitems[i], macro = id, root = parent.root + "/" + parent.name, folder = !fileitems[i].ToLower().EndsWith(".ipy"), children = new List<DataTreeViewItem>() };
                                parent.children.Add(child);
                            }

                            parent = child;
                            lev++;
                        }
                    }
                }
            }

            return items;
        }

#endregion

        #region Tree View Context Menu

        public ContextMenu CreateTreeViewContextMenu()
        {
            Style ContextMenuStyle = MainWindow.GetInstance().GetResource("MetroContextMenuStyle") as Style;
            Style MenuItemStyle = MainWindow.GetInstance().GetResource("MetroMenuItemStyle") as Style;

            ContextMenu cm = new ContextMenu();
            cm.Resources.MergedDictionaries.Add(MainWindow.GetInstance().GetResources());
            cm.Style = ContextMenuStyle;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateMacro(null, MacroType.PYTHON, "/");
                cm.IsOpen = false;
            };
            mi_create.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateFolder(null, "/");
                cm.IsOpen = false;
            };
            mi_folder.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                ImportMacro(null, "/");
                cm.IsOpen = false;
            };
            mi_import.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_import);
            
            return cm;
        }

#endregion

        #region Tree View Item Context Menu

        private ContextMenu CreateContextMenuFolder(DisplayableTreeViewItem item, string name, string path)
        {
            Style ContextMenuStyle = MainWindow.GetInstance().GetResource("MetroContextMenuStyle") as Style;
            Style MenuItemStyle = MainWindow.GetInstance().GetResource("MetroMenuItemStyle") as Style;
            Style SeparatorStyle = MainWindow.GetInstance().GetResource("MertoMenuSeparatorStyle") as Style;

            ContextMenu cm = new ContextMenu();
            cm.Resources.MergedDictionaries.Add(MainWindow.GetInstance().GetResources());
            cm.Style = ContextMenuStyle;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateMacro(item, MacroType.PYTHON, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_create.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateFolder(item, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_folder.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                ImportMacro(item, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_import.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_import);

            Separator sep1 = new Separator();
            sep1.Style = SeparatorStyle;
            cm.Items.Add(sep1);

            MenuItem mi_del = new MenuItem();
            mi_del.Header = "Delete";
            mi_del.Click += delegate (object sender, RoutedEventArgs args)
            {
                DeleteFolder(item, path, name);

                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_del.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                DisplayableTreeViewItem parentitem = item.Parent;
                string previousname = item.Header;

                item.KeyUpEvent = delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        Focus();
                        Keyboard.ClearFocus();
                    }
                    else if (a.Key == Key.Escape)
                    {
                        item.Header = previousname;
                        Focus();
                        Keyboard.ClearFocus();
                    }
                };

                item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
                {
                    if (item.Header == previousname)
                    {
                        item.IsInputting = false;
                        return;
                    }

                    if (String.IsNullOrEmpty(item.Header))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        item.IsInputting = true;
                        return;
                    }
                    
                    MainWindow.GetInstance().RenameFolder(path + name, path + item.Header);
                    Rename(parentitem, item);

                    item.IsInputting = false;
                };

                item.IsInputting = true;
            };
            
            mi_rename.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_rename);

            return cm;
        }

        private ContextMenu CreateContextMenuMacro(DisplayableTreeViewItem item, string name, Guid id)
        {
            Style ContextMenuStyle = MainWindow.GetInstance().GetResource("MetroContextMenuStyle") as Style;
            Style MenuItemStyle = MainWindow.GetInstance().GetResource("MetroMenuItemStyle") as Style;
            Style SeparatorStyle = MainWindow.GetInstance().GetResource("MertoMenuSeparatorStyle") as Style;

            ContextMenu cm = new ContextMenu();
            cm.Resources.MergedDictionaries.Add(MainWindow.GetInstance().GetResources());
            cm.Style = ContextMenuStyle;

            DisplayableTreeViewItem parentitem = item.Parent;

            if (Main.GetMacro(id) == null)
            {
                Routing.EventManager.DisplayOkMessage("Could not find the macro (when attempting to create a context menu): " + name, "Macro Error");
                return null;
            }

            IMacro macro = Main.GetMacro(id);

            MenuItem mi_edit = new MenuItem();
            mi_edit.Header = "Edit";
            mi_edit.Click += delegate (object sender, RoutedEventArgs args)
            {
                OpenMacro(id);
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_edit.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_edit);

            MenuItem mi_execute = new MenuItem();
            mi_execute.Header = "Synchronous Execute";
            mi_execute.ToolTip = "Synchronous executions cannot be terminated.";
            mi_execute.Click += delegate (object sender, RoutedEventArgs args)
            {
                ExecuteMacro(id, macro, false);
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_execute.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_execute);

            MenuItem mi_executea = new MenuItem();
            mi_executea.Header = "Asynchronous Execute";
            mi_executea.ToolTip = "Asynchronous executions can be terminated.";
            mi_executea.Click += delegate (object sender, RoutedEventArgs args)
            {
                ExecuteMacro(id, macro, true);
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_executea.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_executea);

            Separator sep1 = new Separator();
            sep1.Style = SeparatorStyle as Style;
            cm.Items.Add(sep1);

            MenuItem mi_export = new MenuItem();
            mi_export.Header = "Export";
            mi_export.Click += delegate (object sender, RoutedEventArgs args)
            {
                macro.Export();
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_export.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_export);

            MenuItem mi_del = new MenuItem();
            mi_del.Header = "Delete";
            mi_del.Click += delegate (object sender, RoutedEventArgs args)
            {
                DeleteMacro(item, macro);

                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_del.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                string previousname = item.Header;

                item.KeyUpEvent = delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        Focus();
                        Keyboard.ClearFocus();
                    }
                    else if (a.Key == Key.Escape)
                    {
                        item.Header = previousname;
                        Focus();
                        Keyboard.ClearFocus();
                    }
                };

                item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
                {
                    if (item.Header == previousname)
                    {
                        item.IsInputting = false;
                        return;
                    }

                    if (String.IsNullOrEmpty(item.Header))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        item.IsInputting = true;
                        return;
                    }

                    if (!item.Header.EndsWith(FileManager.PYTHON_FILE_EXT))
                        item.Header += FileManager.PYTHON_FILE_EXT;

                    
                    MainWindow.GetInstance().RenameMacro(id, item.Header);Rename(parentitem, item);
                    Rename(parentitem, item);

                    item.IsInputting = false;
                };

                item.IsInputting = true;
            };
            mi_rename.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_rename);

            Separator sep2 = new Separator();
            sep2.Style = SeparatorStyle;
            cm.Items.Add(sep2);

            MenuItem mi_add = new MenuItem();

            mi_add.Click += delegate (object sender, RoutedEventArgs args)
            {
                if (Main.IsRibbonMacro(id))
                    item.IsRibbonMacro = false;
                else
                    item.IsRibbonMacro = true;

                UpdateRibbon();

                mi_add.Header = Main.IsRibbonMacro(id) ? "Remove From Ribbon" : "Add To Ribbon";
                args.Handled = true;
                cm.IsOpen = false;
            };

            mi_add.Header = Main.IsRibbonMacro(id) ? "Remove From Ribbon" : "Add To Ribbon";
            mi_add.Style = MenuItemStyle as Style;
            cm.Items.Add(mi_add);

            return cm;
        }

        private void TreeViewItem_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs args)
        {
            TreeViewItem tvi = sender as TreeViewItem;
            DisplayableTreeViewItem item = tvi.DataContext as DisplayableTreeViewItem;

            SelectedItem = item;

            ContextMenu cm;

            if (!item.IsFolder)
            {
                cm = CreateContextMenuMacro(item, item.Header, item.ID);
            }
            else
            {
                cm = CreateContextMenuFolder(item, item.Header, item.Root);
            }

            if (cm == null)
            {
                System.Diagnostics.Debug.WriteLine("Context menu is null.");
                return;
            }

            tvi.ContextMenu = cm;
            tvi.ContextMenu.IsOpen = true;
        }

#endregion

        #region Tree View Functions & Item Functions
        
        private void CloseItemMacro(DisplayableTreeViewItem item)
        {
            foreach (DisplayableTreeViewItem child in item.Items)
                CloseItemMacro(child);

            MainWindow.GetInstance().CloseMacro(item.ID);
        }

        public void DeleteFolder(DisplayableTreeViewItem item, string path, string name)
        {
            Main.DeleteFolder(path + "/" + name, (result) =>
            {
                if (result)
                {
                    if (item.Parent is DisplayableTreeViewItem)
                        Remove((item.Parent as DisplayableTreeViewItem), item);
                    else
                        Remove(null, item);

                    CloseItemMacro(item);
                }
            });
        }

        public void DeleteMacro(DisplayableTreeViewItem item, IMacro macro)
        {
            macro.Delete((result) =>
            {
                if (result)
                {
                    if (item.Parent is DisplayableTreeViewItem)
                        Remove((item.Parent as DisplayableTreeViewItem), item);
                    else
                        Remove(null, item);

                    CloseItemMacro(item);
                }
            });
        }

        public void ImportMacro()
        {
            if (SelectedItem == null)
                ImportMacro(null, "/");
            else if (SelectedItem.IsFolder)
                ImportMacro(SelectedItem, SelectedItem.Root + '/' + SelectedItem.Header);
            else if (!SelectedItem.IsFolder)
                ImportMacro(SelectedItem.Parent, SelectedItem.Root);
        }

        public void ImportMacro(DisplayableTreeViewItem parent, string relativepath)
        {
            MainWindow.GetInstance().ImportMacro(relativepath, (id) =>
            {
                if (id == Guid.Empty)
                    return;

                DisplayableTreeViewItem item = new DisplayableTreeViewItem();
                item.Header = Main.GetDeclaration(id).name;
                item.IsFolder = false;
                item.ID = id;

                item.RightClickEvent = TreeViewItem_OnPreviewMouseRightButtonDown;

                item.SelectedEvent = delegate (object sender, RoutedEventArgs args) { SelectedItem = item; };
                item.DoubleClickEvent = delegate (object sender, MouseButtonEventArgs args) { OpenMacro(id); args.Handled = true; };

                if (parent == null)
                    item.Root = "/";
                else
                    item.Root = parent.Root + "/" + parent.Header + "/";

                Add(parent, item);
                
                OpenMacro(id);
            });
        }

        public void CreateMacro(MacroType type)
        {
            if (SelectedItem == null)
                CreateMacro(null, type, "/");
            else if(SelectedItem.IsFolder)
                CreateMacro(SelectedItem, type, SelectedItem.Root + '/' + SelectedItem.Header);
            else if (!SelectedItem.IsFolder)
                CreateMacro(SelectedItem.Parent, type, SelectedItem.Root);
        }

        public void CreateMacro(DisplayableTreeViewItem parent, MacroType type, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            DisplayableTreeViewItem item = new DisplayableTreeViewItem();

            item.KeyUpEvent = delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    Focus();
                    Keyboard.ClearFocus();
                }
                else if (a.Key == Key.Escape)
                {
                    m_IsCreating = false;
                    item.IsInputting = false;
                }
            };

            item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
            {
                if (String.IsNullOrEmpty(item.Header) || (!item.IsInputting))
                {
                    Remove(parent, item);
                    m_IsCreating = false;
                    return;
                }

                if (!item.Header.EndsWith(FileManager.PYTHON_FILE_EXT))
                    item.Header += FileManager.PYTHON_FILE_EXT;

                item.Header = Regex.Replace(item.Header, "[^0-9a-zA-Z ._-]", "");

                Guid id = MainWindow.GetInstance().CreateMacro(type, root + "/" + item.Header);
                Rename(parent, item);

                if (id == Guid.Empty)
                {
                    Remove(parent, item);
                    return;
                }

                item.Header = Main.GetDeclaration(id).name;
                item.ID = id;
                item.Root = root;
                item.Parent = parent;
                item.IsFolder = false;

                item.RightClickEvent = TreeViewItem_OnPreviewMouseRightButtonDown;

                item.SelectedEvent = delegate (object sender, RoutedEventArgs args) { SelectedItem = item; };
                item.DoubleClickEvent = delegate (object sender, MouseButtonEventArgs args) { OpenMacro(id); args.Handled = true; };

                item.IsInputting = false;

                m_IsCreating = false;
                OpenMacro(id);
            };

            Add(parent, item);

            item.IsInputting = true;
        }

        private void CreateFolder(DisplayableTreeViewItem parent, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            DisplayableTreeViewItem item = new DisplayableTreeViewItem();

            item.KeyUpEvent = delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    Focus();
                    Keyboard.ClearFocus();
                }
                else if (a.Key == Key.Escape)
                {
                    m_IsCreating = false;
                    item.IsInputting = false;
                }
            };

            item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
            {
                if (String.IsNullOrEmpty(item.Header) || (!item.IsInputting))
                {
                    Remove(parent, item);
                    m_IsCreating = false;
                    return;
                }

                item.Header = Regex.Replace(item.Header, @"[^0-9a-zA-Z]+", "");


                if (!FileManager.CreateFolder((root + "/" + item.Header).Replace('\\', '/').Replace("//", "/")))
                {
                    Remove(parent, item);
                    return;
                }
                Rename(parent, item);

                item.RightClickEvent = TreeViewItem_OnPreviewMouseRightButtonDown;
                
                item.Root = root;
                item.IsFolder = true;
                item.IsExpanded = false;
                item.Parent = parent;

                item.IsInputting = false;

                m_IsCreating = false;
            };

            //tvi.Items.SortDescriptions.Add(new SortDescription("Header", ListSortDirection.Ascending));

            Add(parent, item);

            item.IsInputting = true;
        }

        public void OpenMacro(Guid id)
        {
            MainWindow.GetInstance().OpenMacroForEditing(id);
        }

        public void ExecuteMacro(Guid id, IMacro macro, bool async)
        {
            if (MainWindow.GetInstance().IsActive)
            {
                OpenMacro(id);
                MainWindow.GetInstance().ExecuteMacro(async);
            }
            else
            {
                macro.ExecuteDebug(null, async);
            }
        }

#endregion
    }
}
