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
    /// <summary>
    /// Custom data structure for holding the data of each tree view item
    /// </summary>
    public class CustomTreeViewItem
    {
        public int level;
        public string name;
        public Guid macro;
        public string root;
        public List<CustomTreeViewItem> children;
    }

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

        public CustomTreeViewItem SelectedItem
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

        public ObservableCollection<CustomTreeViewItem> ItemSource
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

        /*#region Tree View Population Through Recursion

        private void CheckVisibility()
        {
            Dictionary<Guid, IMacro>.KeyCollection keys = Main.GetMacros().Keys;
            if (keys.Count == 0)
                lblNoMacros.Visibility = Visibility.Visible;
            else
                lblNoMacros.Visibility = Visibility.Hidden;
        }

        private void Initialize()
        {
            if (((ExplorerViewModel)DataContext).ItemSource == null)
            {
                Dictionary<Guid, IMacro>.KeyCollection keys = Main.GetMacros().Keys;
                HashSet<CustomTreeViewItem> items = CreateTreeViewItemStructure(keys.ToList<Guid>());

                foreach (CustomTreeViewItem item in items)
                {
                    tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
                    {
                        TreeViewItem tvi = CreateTreeViewItem(null, item);

                        if (tvi != null)
                            tvMacroView.Items.Add(tvi);
                    }));
                }

                tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => tvMacroView.Items.SortDescriptions.Add(new SortDescription("Header", ListSortDirection.Ascending))));
                tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => tvMacroView.Items.Refresh()));

                ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;
            }
            else
            {
                foreach (TreeViewItem tvi in ((ExplorerViewModel)DataContext).ItemSource)
                    tvMacroView.Items.Add(AddEvent(tvi));
            }

            CheckVisibility();
            Main.GetInstance().OnMacroCountChanged += CheckVisibility;
        }

        private TreeViewItem AddEvent(TreeViewItem item)
        {
            if (item == null)
                return item;

            item.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;

            for (int i = 0; i < item.Items.Count; i++)
                item.Items[i] = AddEvent(item.Items[i] as TreeViewItem);

            return item;
        }

        private TreeViewItem CreateTreeViewItem(TreeViewItem parent, CustomTreeViewItem ctvi)
        {
            TreeViewItem tvi = new TreeViewItem();
            tvi.DataContext = ctvi;
            tvi.Header = ctvi.name;

            if (ctvi.children != null)
            {
                foreach (CustomTreeViewItem child in ctvi.children)
                {
                    TreeViewItem node = CreateTreeViewItem(tvi, child);

                    if (node != null)
                        tvMacroView.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => tvi.Items.Add(node)));
                }
            }

            if (ctvi.name.ToLower().Trim().EndsWith(FileManager.PYTHON_FILE_EXT))
            {
                tvi.Tag = ctvi.macro.ToString();
                tvi.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;

                tvi.Selected += delegate (object sender, RoutedEventArgs args) { ((ExplorerViewModel)DataContext).SelectedItem = tvi; };
                tvi.MouseDoubleClick += delegate (object sender, MouseButtonEventArgs args) { OpenMacro(ctvi.macro); args.Handled = true; };
            }
            else
            {
                tvi.Tag = ctvi.root;
                tvi.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;
            }

            return tvi;
        }

        private HashSet<CustomTreeViewItem> CreateTreeViewItemStructure(List<Guid> macros)
        {
            HashSet<CustomTreeViewItem> items = new HashSet<CustomTreeViewItem>();

            foreach (Guid id in macros)
            {
                string path = tvMacroView.Dispatcher.Invoke(() => Regex.Replace(Main.GetDeclaration(id).relativepath, @"/+", System.IO.Path.DirectorySeparatorChar.ToString()));
                string[] fileitems = path.Split(System.IO.Path.DirectorySeparatorChar).Where<string>(x => !String.IsNullOrEmpty(x)).ToArray<string>();

                if (fileitems.Any())
                {
                    CustomTreeViewItem root = items.FirstOrDefault(x => x.name.Equals(fileitems[0]) && x.level.Equals(1));

                    if (root == null)
                    {
                        root = new CustomTreeViewItem() { level = 1, name = fileitems[0], macro = id, root = "/", children = new List<CustomTreeViewItem>() };
                        items.Add(root);
                    }

                    if (fileitems.Length > 1)
                    {
                        CustomTreeViewItem parent = root;
                        int lev = 2;

                        for (int i = 1; i < fileitems.Length; i++)
                        {
                            CustomTreeViewItem child = parent.children.FirstOrDefault(x => x.name.Equals(fileitems[i]) && x.level.Equals(lev));

                            if (child == null)
                            {
                                child = new CustomTreeViewItem() { level = lev, name = fileitems[i], macro = id, root = parent.root + "/" + parent.name, children = new List<CustomTreeViewItem>() };
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

        private void ExplorerView_Loaded(object s, RoutedEventArgs e)
        {
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateMacro(null, MacroType.PYTHON, "/");
                tvMacroView.ContextMenu.IsOpen = false;
            };
            mi_create.DataContext = tvMacroView;
            mi_create.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                CreateFolder(null, "/");
                tvMacroView.ContextMenu.IsOpen = false;
            };
            mi_folder.DataContext = tvMacroView;
            mi_folder.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                ImportMacro(null, "/");
                tvMacroView.ContextMenu.IsOpen = false;
            };
            mi_import.DataContext = tvMacroView;
            mi_import.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_import);

            tvMacroView.ContextMenu = cm;
            tvMacroView.MouseRightButtonDown += delegate (object sender, MouseButtonEventArgs args) { tvMacroView.ContextMenu.IsOpen = true; args.Handled = true; };

            Initialize();
        }

        #endregion

        #region Tree View Item Context Menu

        private ContextMenu CreateContextMenuFolder(TreeViewItem item, string name, string path)
        {
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

            MenuItem mi_create = new MenuItem();
            mi_create.Header = "Create Macro";
            mi_create.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateMacro(item, MacroType.PYTHON, path + "/" + name + "/");
                item.ContextMenu.IsOpen = false;
            };
            mi_create.DataContext = item;
            mi_create.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_create);

            MenuItem mi_folder = new MenuItem();
            mi_folder.Header = "Create Folder";
            mi_folder.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                CreateFolder(item, path + "/" + name + "/");
                item.ContextMenu.IsOpen = false;
            };
            mi_folder.DataContext = item;
            mi_folder.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_folder);

            MenuItem mi_import = new MenuItem();
            mi_import.Header = "Import Macro";
            mi_import.Click += delegate (object sender, RoutedEventArgs args)
            {
                item.IsExpanded = true;
                ImportMacro(item, path + "/" + name + "/");
                item.ContextMenu.IsOpen = false;
            };
            mi_import.DataContext = item;
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
            mi_del.DataContext = item;
            mi_del.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                string previousname = item.Header.ToString();

                TextBox inputBox = new TextBox();
                inputBox.BorderThickness = new Thickness(1);

                inputBox.KeyUp += delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { item.Focus(); Keyboard.ClearFocus(); }));
                    }
                    else if (a.Key == Key.Escape)
                    {
                        inputBox.Visibility = Visibility.Hidden;
                        item.Header = previousname;
                    }
                };

                inputBox.LostFocus += delegate (object s, RoutedEventArgs a)
                {
                    if (!inputBox.IsVisible)
                        return;

                    if (item.Header.GetType() != typeof(TextBox))
                    {
                        Routing.EventManager.DisplayOkMessage("Error renaming the folder.", "Renaming Error");
                        if (item.Parent is TreeViewItem)
                            (item.Parent as TreeViewItem).Items.Remove(item);
                        else
                            tvMacroView.Items.Remove(item);
                        return;
                    }
                    else if (String.IsNullOrEmpty((item.Header as TextBox).Text))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
                        return;
                    }

                    int index = -1;

                    DependencyObject parentobj = item.Parent;
                    TreeViewItem parentitem = GetDependencyObjectFromVisualTree(parentobj, typeof(TreeViewItem)) as TreeViewItem;

                    if (parentitem == null)
                        index = tvMacroView.Items.IndexOf(item);
                    else
                        index = parentitem.Items.IndexOf(item);

                    if (index == -1)
                        return;

                    string newname = (item.Header as TextBox).Text;
                    Dispatcher.Invoke(() => Main.RenameFolder(path + name, path + newname));
                    item.Header = (item.Header as TextBox).Text;

                    ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;
                };

                item.Header = inputBox;
                Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
            };

            mi_rename.DataContext = item;
            mi_rename.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_rename);

            return cm;
        }

        private ContextMenu CreateContextMenuMacro(TreeViewItem item, string name, Guid id)
        {
            ContextMenu cm = new ContextMenu();
            cm.Style = Resources["ContextMenuMetroStyle"] as Style;

            DependencyObject parentobj = item.Parent;
            TreeViewItem parentitem = GetDependencyObjectFromVisualTree(parentobj, typeof(TreeViewItem)) as TreeViewItem;

            if (!Dispatcher.Invoke(() => Main.GetMacros().ContainsKey(id)))
            {
                Routing.EventManager.DisplayOkMessage("Could not find the macro (when attempting to create a context menu): " + name, "Macro Error");
                return null;
            }

            IMacro macro = Dispatcher.Invoke(() => Main.GetMacros()[id]);

            MenuItem mi_edit = new MenuItem();
            mi_edit.Header = "Edit";
            mi_edit.Click += delegate (object sender, RoutedEventArgs args)
            {
                OpenMacro(id);
                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_edit.DataContext = item;
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
            mi_execute.DataContext = item;
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
            mi_executea.DataContext = item;
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
            mi_export.DataContext = item;
            mi_export.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_export);

            MenuItem mi_del = new MenuItem();
            mi_del.Header = "Delete";
            mi_del.Click += delegate (object sender, RoutedEventArgs args)
            {
                DeleteMacro(parentitem, item, macro);

                args.Handled = true;
                cm.IsOpen = false;
            };
            mi_del.DataContext = item;
            mi_del.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_del);

            MenuItem mi_rename = new MenuItem();
            mi_rename.Header = "Rename";
            mi_rename.Click += delegate (object sender, RoutedEventArgs args)
            {
                args.Handled = true;
                cm.IsOpen = false;

                string previousname = item.Header.ToString();

                TextBox inputBox = new TextBox();
                inputBox.BorderThickness = new Thickness(1);

                inputBox.KeyUp += delegate (object s, KeyEventArgs a)
                {
                    if (a.Key == Key.Return)
                    {
                        Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { item.Focus(); Keyboard.ClearFocus(); }));
                    }
                    else if (a.Key == Key.Escape)
                    {
                        inputBox.Visibility = Visibility.Hidden;
                        item.Header = previousname;
                    }
                };

                inputBox.LostFocus += delegate (object s, RoutedEventArgs a)
                {
                    if (!inputBox.IsVisible)
                        return;

                    if (item.Header.GetType() != typeof(TextBox))
                    {
                        Routing.EventManager.DisplayOkMessage("Error renaming the macro.", "Renaming Error");
                        tvMacroView.Items.Remove(item);
                        return;
                    }
                    else if (String.IsNullOrEmpty((item.Header as TextBox).Text))
                    {
                        Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                        Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
                        return;
                    }

                    if (!(item.Header as TextBox).Text.EndsWith(FileManager.PYTHON_FILE_EXT))
                        (item.Header as TextBox).Text += FileManager.PYTHON_FILE_EXT;

                    int index = -1;

                    if (parentitem == null)
                    {

                        index = tvMacroView.Items.IndexOf(item);

                        if (index == -1)
                            return;

                        string newname = (item.Header as TextBox).Text;
                        Guid decl = Dispatcher.Invoke(() => Main.RenameMacro(id, newname));
                        (tvMacroView.Items[index] as TreeViewItem).Header = Dispatcher.Invoke(() => Main.GetDeclaration(decl).name);
                    }
                    else
                    {
                        index = parentitem.Items.IndexOf(item);

                        if (index == -1)
                            return;

                        string newname = (item.Header as TextBox).Text;
                        Guid decl = Dispatcher.Invoke(() => Main.RenameMacro(id, newname));
                        (parentitem.Items[index] as TreeViewItem).Header = Dispatcher.Invoke(() => Main.GetDeclaration(decl).name);
                    }

                    ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;
                };

                item.Header = inputBox;
                Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
            };
            mi_rename.DataContext = item;
            mi_rename.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_rename);

            Separator sep2 = new Separator();
            sep2.Style = Resources["MenuSeparatorMertoStyle"] as Style;
            cm.Items.Add(sep2);

            MenuItem mi_add = new MenuItem();

            mi_add.Click += delegate (object sender, RoutedEventArgs args)
            {
                if (Dispatcher.Invoke(() => Main.IsRibbonMacro(id)))
                    Main.AddRibbonMacro(id);
                else
                    Main.RemoveRibbonMacro(id);

                mi_add.Header = Dispatcher.Invoke(() => Main.IsRibbonMacro(id)) ? "Remove From Ribbon" : "Add To Ribbon";
                args.Handled = true;
                cm.IsOpen = false;
            };

            mi_add.Header = Dispatcher.Invoke(() => Main.IsRibbonMacro(id)) ? "Remove From Ribbon" : "Add To Ribbon";
            mi_add.DataContext = item;
            mi_add.Style = Resources["MenuItemMetroStyle"] as Style;
            cm.Items.Add(mi_add);

            return cm;
        }

        private void TreeViewItem_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs args)
        {
            DependencyObject obj = args.OriginalSource as DependencyObject;
            TreeViewItem item = GetDependencyObjectFromVisualTree(obj, typeof(TreeViewItem)) as TreeViewItem;

            item.IsSelected = true;

            string header = item.Header.ToString();
            string tagstring = item.Tag as string;

            ContextMenu cm;
            //ThemeManager.ChangeAppStyle(Main.GetExcelDispatcher().Invoke(() => Main.GetMacroManager()), ThemeManager.GetAccent("ExcelAccent"), ThemeManager.GetAppTheme("BaseLight"));

            if (header.ToLower().Trim().EndsWith(FileManager.PYTHON_FILE_EXT))
            {
                cm = CreateContextMenuMacro(item, header, Guid.Parse(tagstring));
            }
            else
            {
                cm = CreateContextMenuFolder(item, header, tagstring);
            }

            if (cm == null)
            {
                System.Diagnostics.Debug.WriteLine("Context menu is null.");
                return;
            }

            item.ContextMenu = cm;
            item.ContextMenu.IsOpen = true;
        }

        #endregion

        #region Tree View Functions & Item Functions

        private DependencyObject GetDependencyObjectFromVisualTree(DependencyObject obj, Type type)
        {
            DependencyObject parent = obj;

            while (parent != null)
            {
                if (type.IsInstanceOfType(parent))
                    break;

                parent = VisualTreeHelper.GetParent(parent);
            }

            return parent;
        }

        public void DeleteFolder(TreeViewItem item, string path, string name)
        {
            Main.DeleteFolder(path + "/" + name, async (result) =>
            {
                if (result)
                {
                    if (item.Parent is TreeViewItem)
                        await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { (item.Parent as TreeViewItem).Items.Remove(item); }));
                    else
                        await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { tvMacroView.Items.Remove(item); }));

                    await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(() => ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items));
                }
            });
        }

        public void DeleteMacro(TreeViewItem parent, TreeViewItem item, IMacro macro)
        {
            macro.Delete(async (result) =>
            {
                if (result)
                {
                    if (item.Parent is TreeViewItem)
                        await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { (item.Parent as TreeViewItem).Items.Remove(item); }));
                    else
                        await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { tvMacroView.Items.Remove(item); }));

                    await Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(() => ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items));
                }
            });
        }

        public void ImportMacro(TreeViewItem parent, string relativepath)
        {
            MainWindow.GetInstance().ImportMacro(relativepath, (id) =>
            {
                if (id == Guid.Empty)
                    return;

                TreeViewItem tvi = new TreeViewItem();
                tvi.Header = tvMacroView.Dispatcher.Invoke(() => Main.GetDeclaration(id).name);
                tvi.Tag = id.ToString(); ;

                tvi.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;

                tvi.Selected += delegate (object sender, RoutedEventArgs args) { ((ExplorerViewModel)DataContext).SelectedItem = tvi; };
                tvi.MouseDoubleClick += delegate (object sender, MouseButtonEventArgs args) { OpenMacro(id); args.Handled = true; };

                if (parent == null)
                    tvMacroView.Items.Add(tvi);
                else
                    parent.Items.Add(tvi);

                ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;

                OpenMacro(id);
            });
        }

        public void CreateMacro(TreeViewItem parent, MacroType type, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            TextBox inputBox = new TextBox();
            inputBox.BorderThickness = new Thickness(1);

            TreeViewItem tvi = new TreeViewItem();

            inputBox.KeyUp += delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { tvi.Focus(); Keyboard.ClearFocus(); }));
                }
                else if (a.Key == Key.Escape)
                {
                    inputBox.Visibility = Visibility.Hidden;

                    if (parent == null)
                        tvMacroView.Items.Remove(tvi);
                    else
                        parent.Items.Remove(tvi);
                }
            };

            inputBox.LostFocus += delegate (object s, RoutedEventArgs a)
            {
                if (!inputBox.IsVisible)
                    return;

                if (tvi.Header.GetType() != typeof(TextBox))
                {
                    Routing.EventManager.DisplayOkMessage("Error creating the macro.", "Creation error");
                    tvMacroView.Items.Remove(tvi);
                    return;
                }
                else if (String.IsNullOrEmpty((tvi.Header as TextBox).Text))
                {
                    Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid name");
                    tvMacroView.Items.Remove(tvi);
                    CreateMacro(parent, type, root);
                    return;
                }

                if (!(tvi.Header as TextBox).Text.EndsWith(FileManager.PYTHON_FILE_EXT))
                    (tvi.Header as TextBox).Text += FileManager.PYTHON_FILE_EXT;

                (tvi.Header as TextBox).Text = Regex.Replace((tvi.Header as TextBox).Text, "[^0-9a-zA-Z ._-]", "");

                Guid id = MainWindow.GetInstance().CreateMacro(type, root + "/" + (tvi.Header as TextBox).Text);

                if (id == Guid.Empty)
                {
                    tvMacroView.Items.Remove(tvi);

                    if (parent == null)
                        tvMacroView.Items.Refresh();
                    else
                        parent.Items.Refresh();

                    return;
                }

                tvi.Header = tvMacroView.Dispatcher.Invoke(() => Main.GetDeclaration(id).name);
                tvi.DataContext = new CustomTreeViewItem() { name = Main.GetDeclaration(id).name, macro = id, root = root };
                tvi.Tag = id.ToString();

                if (parent == null)
                    tvMacroView.Items.Refresh();
                else
                    parent.Items.Refresh();

                tvi.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;

                tvi.Selected += delegate (object sender, RoutedEventArgs args) { ((ExplorerViewModel)DataContext).SelectedItem = tvi; };
                tvi.MouseDoubleClick += delegate (object sender, MouseButtonEventArgs args) { OpenMacro(id); args.Handled = true; };

                ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;

                m_IsCreating = false;
                OpenMacro(id);
            };

            tvi.Header = inputBox;

            if (parent == null)
                tvMacroView.Items.Add(tvi);
            else
                parent.Items.Add(tvi);

            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
        }

        private void CreateFolder(TreeViewItem parent, string root)
        {
            if (m_IsCreating)
                return;

            m_IsCreating = true;

            TextBox inputBox = new TextBox();
            inputBox.BorderThickness = new Thickness(1);

            TreeViewItem tvi = new TreeViewItem();

            inputBox.KeyUp += delegate (object s, KeyEventArgs a)
            {
                if (a.Key == Key.Return)
                {
                    Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { tvi.Focus(); Keyboard.ClearFocus(); }));
                }
                else if (a.Key == Key.Escape)
                {
                    inputBox.Visibility = Visibility.Hidden;

                    if (parent == null)
                        tvMacroView.Items.Remove(tvi);
                    else
                        parent.Items.Remove(tvi);
                }
            };

            inputBox.LostFocus += delegate (object s, RoutedEventArgs a)
            {
                if (tvi.Header.GetType() != typeof(TextBox))
                {
                    Routing.EventManager.DisplayOkMessage("Error creating the folder.", "Creation Error");
                    tvMacroView.Items.Remove(tvi);
                    return;
                }
                else if (String.IsNullOrEmpty((tvi.Header as TextBox).Text))
                {
                    Routing.EventManager.DisplayOkMessage("Please enter a valid name.", "Invalid Name");
                    CreateFolder(parent, root);
                    return;
                }

                (tvi.Header as TextBox).Text = Regex.Replace((tvi.Header as TextBox).Text, @"[^0-9a-zA-Z]+", "");

                string newname = (tvi.Header as TextBox).Text;
                Dispatcher.Invoke(() => FileManager.CreateFolder((root + "/" + newname).Replace('\\', '/').Replace("//", "/")));

                tvi.PreviewMouseRightButtonDown += TreeViewItem_OnPreviewMouseRightButtonDown;

                tvi.Header = (tvi.Header as TextBox).Text;

                if (parent == null)
                    tvMacroView.Items.Refresh();
                else
                    parent.Items.Refresh();

                ((ExplorerViewModel)DataContext).ItemSource = tvMacroView.Items;

                m_IsCreating = false;
            };

            tvi.Header = inputBox;
            tvi.Tag = root;
            tvi.Items.SortDescriptions.Add(new SortDescription("Header", ListSortDirection.Ascending));

            if (parent == null)
                tvMacroView.Items.Add(tvi);
            else
                parent.Items.Add(tvi);

            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { inputBox.Focus(); }));
        }

        public void OpenMacro(Guid id)
        {
            Dispatcher.Invoke(() => MainWindow.GetInstance().OpenMacroForEditing(id));
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

        #endregion*/
    }
}
