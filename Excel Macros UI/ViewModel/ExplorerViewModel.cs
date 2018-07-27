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

        #region Tree View Population Through Recursion

        private void CheckVisibility()
        {
            Dictionary<Guid, IMacro>.KeyCollection keys = Main.GetMacros().Keys;
            if (keys.Count == 0)
                LabelVisibility = Visibility.Visible;
            else
                LabelVisibility = Visibility.Hidden;
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

            //tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => tvMacroView.Items.SortDescriptions.Add(new SortDescription("Header", ListSortDirection.Ascending))));
            //tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => tvMacroView.Items.Refresh()));

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
                        root = new DataTreeViewItem() { level = 1, name = fileitems[0], macro = id, root = "/", folder = true, children = new List<DataTreeViewItem>() };
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
                                child = new DataTreeViewItem() { level = lev, name = fileitems[i], macro = id, root = parent.root + "/" + parent.name, folder = !(fileitems[i].ToLower().EndsWith(".ipy") || fileitems[i].ToLower().EndsWith(".ipy")), children = new List<DataTreeViewItem>() };
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
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateMacro(null, MacroType.PYTHON, "/");
                cm.IsOpen = false;
            };
            mi_create.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateFolder(null, "/");
                cm.IsOpen = false;
            };
            mi_folder.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                ImportMacro(null, "/");
                cm.IsOpen = false;
            };
            mi_import.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_import);
            
            return cm;
        }

        #endregion

        #region Tree View Item Context Menu

        private ContextMenu CreateContextMenuFolder(DisplayableTreeViewItem item, string name, string path)
        {
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateMacro(item, MacroType.PYTHON, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_create.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateFolder(item, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_folder.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                ImportMacro(item, path + "/" + name + "/");
                cm.IsOpen = false;
            };
            mi_import.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_import);

            Separator sep1 = new Separator();
            sep1.Style = Resources["MenuSeparatorMertoStyle"] as Style;
            cm.Items.Add(sep1);

            MenuItem mi_del = new MenuItem();
            mi_del.Header = "Delete";
            mi_del.Click += delegate (object sender, RoutedEventArgs args)
            {
                DeleteFolder(item, path, name);

                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_del.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                DisplayableTreeViewItem parentitem = item.Parent;

                string previousname = item.Header;
                int index = -1;

                if (parentitem == null)
                    index = ItemSource.IndexOf(item);
                else
                    index = parentitem.Items.IndexOf(item);

                if (index == -1)
                    return;

                //TextBox inputBox = new TextBox();
                //inputBox.BorderThickness = new Thickness(1);

                InputableTreeViewItem inputitem = new InputableTreeViewItem();

                inputitem.KeyUpEvent = delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        MainWindow.GetInstance().Focus();
                        Keyboard.ClearFocus();
                    }
                    else if (a.Key == Key.Escape)
                    {
                        inputitem.Visibility = Visibility.Hidden;
                        item.Header = previousname;
                    }
                };

                inputitem.FocusLostEvent = delegate (object s, RoutedEventArgs a)
                {
                    if (inputitem.Visibility != Visibility.Visible)
                        return;
                    
                    if (String.IsNullOrEmpty(inputitem.Header))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        inputitem.FocusEvent?.BeginInvoke(null, null);
                        return;
                    }

                    string newname = inputitem.Header;
                    Main.RenameFolder(path + name, path + newname);
                    item.Header = newname;

                    if (parentitem == null)
                        ItemSource.Remove(inputitem);
                    else
                        parentitem.Items.Remove(inputitem);
                };

                item.Visibility = Visibility.Hidden;

                if (parentitem == null)
                    ItemSource.Add(inputitem);
                else
                    parentitem.Items.Add(inputitem);
            };
            
            mi_rename.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_rename);

            return cm;
        }

        private ContextMenu CreateContextMenuMacro(DisplayableTreeViewItem item, string name, Guid id)
        {
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

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
            mi_edit.Style = Resources["MenuItemMetroStyle"] as Style;
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
            mi_execute.Style = Resources["MenuItemMetroStyle"] as Style;
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
            mi_executea.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_executea);

            Separator sep1 = new Separator();
            sep1.Style = Resources["MenuSeparatorMertoStyle"] as Style;
            cm.Items.Add(sep1);

            MenuItem mi_export = new MenuItem();
            mi_export.Header = "Export";
            mi_export.Click += delegate (object sender, RoutedEventArgs args)
            {
                macro.Export();
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_export.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_export);

            MenuItem mi_del = new MenuItem();
            mi_del.Header = "Delete";
            mi_del.Click += delegate (object sender, RoutedEventArgs args)
            {
                DeleteMacro(item, macro);

                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_del.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                string previousname = item.Header;
                int index = -1;

                if (parentitem == null)
                    index = ItemSource.IndexOf(item);
                else
                    index = parentitem.Items.IndexOf(item);

                if (index == -1)
                    return;

                //TextBox inputBox = new TextBox();
                //inputBox.BorderThickness = new Thickness(1);

                InputableTreeViewItem inputitem = new InputableTreeViewItem();

                inputitem.KeyUpEvent = delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        MainWindow.GetInstance().Focus();
                        Keyboard.ClearFocus();
                    }
                    else if (a.Key == Key.Escape)
                    {
                        inputitem.Visibility = Visibility.Hidden;
                        inputitem.Header = previousname;
                    }
                };

                inputitem.FocusLostEvent = delegate (object s, RoutedEventArgs a)
                {
                    if (inputitem.Visibility != Visibility.Visible)
                        return;

                    
                    if (String.IsNullOrEmpty(inputitem.Header))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        inputitem.FocusEvent?.Invoke();
                        return;
                    }

                    if (!inputitem.Header.EndsWith(FileManager.PYTHON_FILE_EXT))
                        inputitem.Header += FileManager.PYTHON_FILE_EXT;


                    string newname = inputitem.Header;
                    Guid decl = Main.RenameMacro(id, newname);
                    item.Header = newname;
                    item.Visibility = Visibility.Visible;
                    
                    if (parentitem == null)
                        ItemSource.Remove(inputitem);
                    else
                        parentitem.Items.Remove(inputitem);
                };

                item.Visibility = Visibility.Hidden;

                if (parentitem == null)
                    ItemSource.Add(inputitem);
                else
                    parentitem.Items.Add(inputitem);

                inputitem.FocusEvent?.BeginInvoke(null, null);
            };
            mi_rename.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_rename);

            Separator sep2 = new Separator();
            sep2.Style = Resources["MenuSeparatorMertoStyle"] as Style;
            cm.Items.Add(sep2);

            MenuItem mi_add = new MenuItem();

            mi_add.Click += delegate (object sender, RoutedEventArgs args)
            {
                if (Main.IsRibbonMacro(id))
                    Main.AddRibbonMacro(id);
                else
                    Main.RemoveRibbonMacro(id);

                mi_add.Header = Main.IsRibbonMacro(id) ? "Remove From Ribbon" : "Add To Ribbon";
                args.Handled = true;
                cm.IsOpen = false;
            };

            mi_add.Header = Main.IsRibbonMacro(id) ? "Remove From Ribbon" : "Add To Ribbon";
            mi_add.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_add);

            return cm;
        }

        private void TreeViewItem_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs args)
        {
            DependencyObject obj = args.OriginalSource as DependencyObject;
            TreeViewItem tvi = obj as TreeViewItem;
            DisplayableTreeViewItem item = tvi.DataContext as DisplayableTreeViewItem;

            SelectedItem = item;

            ContextMenu cm;
            //ThemeManager.ChangeAppStyle(Main.GetExcelDispatcher().Invoke(() => Main.GetMacroManager()), ThemeManager.GetAccent("ExcelAccent"), ThemeManager.GetAppTheme("BaseLight"));

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
        
        public void DeleteFolder(DisplayableTreeViewItem item, string path, string name)
        {
            Main.DeleteFolder(path + "/" + name, (result) =>
            {
                if (result)
                {
                    if (item.Parent is DisplayableTreeViewItem)
                        (item.Parent as DisplayableTreeViewItem).Items.Remove(item);
                    else
                        ItemSource.Remove(item);
                }
            });
        }

        public void DeleteMacro(DisplayableTreeViewItem item, IMacro macro)
        {
            macro.Delete((result) =>
            {
                if (result)
                {
                    if (item.Parent != null)
                        (item.Parent as DisplayableTreeViewItem).Items.Remove(item);
                    else
                        ItemSource.Remove(item);
                }
            });
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
                {
                    item.Root = "/";
                    ItemSource.Add(item);
                }
                else
                {
                    item.Root = parent.Root + "/" + parent.Header + "/";
                    parent.Items.Add(item);
                }
                
                OpenMacro(id);
            });
        }

        public void CreateMacro(DisplayableTreeViewItem parent, MacroType type, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            //TextBox inputBox = new TextBox();
            //inputBox.BorderThickness = new Thickness(1);

            InputableTreeViewItem item = new InputableTreeViewItem();

            item.KeyUpEvent = delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    MainWindow.GetInstance().Focus();
                    Keyboard.ClearFocus();
                }
                else if (a.Key == Key.Escape)
                {
                    item.Visibility = Visibility.Hidden;

                    if (parent == null)
                        ItemSource.Remove(item);
                    else
                        parent.Items.Remove(item);
                }
            };

            item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
            {
                if (item.Visibility != Visibility.Visible)
                    return;
                
                if (String.IsNullOrEmpty(item.Header))
                {
                    Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid name");

                    if (parent == null)
                        ItemSource.Remove(item);
                    else
                        parent.Items.Remove(item);

                    CreateMacro(parent, type, root);
                    return;
                }

                if (!item.Header.EndsWith(FileManager.PYTHON_FILE_EXT))
                    item.Header += FileManager.PYTHON_FILE_EXT;

                item.Header = Regex.Replace(item.Header, "[^0-9a-zA-Z ._-]", "");

                Guid id = MainWindow.GetInstance().CreateMacro(type, root + "/" + item.Header);
                
                if (parent == null)
                    ItemSource.Remove(item);
                else
                    parent.Items.Remove(item);

                if (id == Guid.Empty)
                    return;

                DisplayableTreeViewItem newItem = new DisplayableTreeViewItem();

                newItem.Header = Main.GetDeclaration(id).name;
                newItem.ID = id;
                newItem.Root = root;
                newItem.Parent = parent;
                newItem.IsFolder = false;

                newItem.RightClickEvent = TreeViewItem_OnPreviewMouseRightButtonDown;

                newItem.SelectedEvent = delegate (object sender, RoutedEventArgs args) { SelectedItem = newItem; };
                newItem.DoubleClickEvent = delegate (object sender, MouseButtonEventArgs args) { OpenMacro(id); args.Handled = true; };
                
                if (parent == null)
                    ItemSource.Add(newItem);
                else
                    parent.Items.Add(newItem);

                m_IsCreating = false;
                OpenMacro(id);
            };
            
            if (parent == null)
                ItemSource.Add(item);
            else
                parent.Items.Add(item);

            item.FocusEvent?.BeginInvoke(null, null);
        }

        private void CreateFolder(DisplayableTreeViewItem parent, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            //TextBox inputBox = new TextBox();
            //inputBox.BorderThickness = new Thickness(1);

            InputableTreeViewItem item = new InputableTreeViewItem();

            item.KeyUpEvent = delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    MainWindow.GetInstance().Focus();
                    Keyboard.ClearFocus();
                }
                else if (a.Key == Key.Escape)
                {
                    item.Visibility = Visibility.Hidden;

                    if (parent == null)
                        ItemSource.Remove(item);
                    else
                        parent.Items.Remove(item);
                }
            };

            item.FocusLostEvent = delegate (object s, RoutedEventArgs a)
            {
                if (String.IsNullOrEmpty(item.Header))
                {
                    Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                    CreateFolder(parent, root);
                    return;
                }

                item.Header = Regex.Replace(item.Header, @"[^0-9a-zA-Z]+", "");

                string newname = item.Header;
                FileManager.CreateFolder((root + "/" + newname).Replace('\\', '/').Replace("//", "/"));

                if (parent == null)
                    ItemSource.Remove(item);
                else
                    parent.Items.Remove(item);

                DisplayableTreeViewItem newItem = new DisplayableTreeViewItem();

                newItem.RightClickEvent = TreeViewItem_OnPreviewMouseRightButtonDown;

                newItem.Header = newname;
                newItem.Root = root;
                newItem.IsFolder = true;
                newItem.IsExpanded = false;
                newItem.Parent = parent;

                if (parent == null)
                    ItemSource.Add(newItem);
                else
                    parent.Items.Add(newItem);

                m_IsCreating = false;
            };
            
            item.Root = root;
            //tvi.Items.SortDescriptions.Add(new SortDescription("Header", ListSortDirection.Ascending));

            if (parent == null)
                ItemSource.Add(item);
            else
                parent.Items.Add(item);

            item.FocusEvent?.BeginInvoke(null, null);
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
