/*
 * Mark Diedericks
 * 24/07/2018
 * Version 1.0.3
 * Handles the view models of the primary view model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class DockManagerViewModel
    {
        public ObservableCollection<DocumentViewModel> Documents { get; internal set; }
        public ObservableCollection<ToolViewModel> Tools { get; internal set; }

        public ExplorerViewModel Explorer { get; internal set; }
        public ToolboxViewModel Toolbox { get; internal set; }
        public ConsoleViewModel Console { get; internal set; }

        public DockManagerViewModel(IEnumerable<DocumentViewModel> DocumentViewModels)
        {
            Documents = new ObservableCollection<DocumentViewModel>();
            Tools = new ObservableCollection<ToolViewModel>();

            Explorer = new ExplorerViewModel() { Model = new ExplorerModel() { Title = "Explorer", ContentId = "Explorer", IsVisible = true } };
            Toolbox = new ToolboxViewModel() { Model = new ToolboxModel() { Title = "Toolbox", ContentId = "Toolbox", IsVisible = true } };
            Console = new ConsoleViewModel() { Model = new ConsoleModel() { Title = "Console", ContentId = "Console", IsVisible = true } };

            Tools.Add(Explorer);
            Tools.Add(Toolbox);
            Tools.Add(Console);

            foreach (DocumentViewModel document in DocumentViewModels)
            {
                document.PropertyChanged += Document_PropertyChanged;

                if (!document.IsClosed)
                    Documents.Add(document);
            }
        }

        public void AddDocument(DocumentViewModel document)
        {
            document.PropertyChanged += Document_PropertyChanged;

            if (!document.IsClosed)
                Documents.Add(document);
        }

        public DockManagerViewModel(string docs) : this(LoadVisibleDocuments(docs))
        {

        }

        private void Document_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            DocumentViewModel document = (DocumentViewModel)sender;

            if(e.PropertyName == nameof(DocumentViewModel.IsClosed))
            {
                if (!document.IsClosed)
                    Documents.Add(document);
                else
                    Documents.Remove(document);
            }
        }

        private static List<DocumentViewModel> LoadVisibleDocuments(string docs)
        {
            string[] paths = docs.Split(';');
            List<DocumentViewModel> documents = new List<DocumentViewModel>();

            foreach(string s in paths)
            {
                Guid id = Main.GetGuidFromRelativePath(s);
                if (id != Guid.Empty)
                {
                    DocumentModel model = DocumentModel.Create(id);

                    if (model != null)
                    {
                        DocumentViewModel viewModel = DocumentViewModel.Create(model);
                        documents.Add(viewModel);
                    }
                }
            }

            return documents;
        }

        public string GetVisibleDocuments()
        {
            StringBuilder sb = new StringBuilder();

            foreach(DocumentViewModel document in Documents)
            {
                if(document.Model.Macro != Guid.Empty)
                {
                    IMacro macro = Main.GetMacro(document.Model.Macro);

                    if(macro != null)
                    {
                        sb.Append(macro.GetRelativePath());
                        sb.Append(';');
                    }
                }
            }

            return sb.ToString();
        }
    }
}
