/*
 * Mark Diedericks
 * 31/07/2018
 * Version 1.0.4
 * Handles the interaction logic of the dock view
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
    public class DockManagerViewModel : Base.ViewModel
    {
        public DockManagerViewModel(IEnumerable<DocumentViewModel> DocumentViewModels)
        {
            Model = new DockManagerModel();

            Explorer = new ExplorerViewModel() { Model = new ExplorerModel() { Title = "Explorer", ContentId = "Explorer", IsVisible = true } };
            Console = new ConsoleViewModel() { Model = new ConsoleModel() { Title = "Console", ContentId = "Console", IsVisible = true } };

            Tools.Add(Explorer);
            Tools.Add(Console);

            {
                VisualEditorViewModel vevm = new VisualEditorViewModel();
                vevm.PropertyChanged += Document_PropertyChanged;

                if (!vevm.IsClosed)
                    Documents.Add(vevm);
            }

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

        public DocumentViewModel GetDocument(Guid id)
        {
            foreach (DocumentViewModel document in Documents)
                if (document.Macro == id)
                    return document;

            return null;
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

        #region Model

        private DockManagerModel m_Model;
        public DockManagerModel Model
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

        #region Documents

        public ObservableCollection<DocumentViewModel> Documents
        {
            get
            {
                return Model.Documents;
            }
            set
            {
                if(Model.Documents != value)
                {
                    Model.Documents = value;
                    OnPropertyChanged(nameof(Documents));
                }
            }
        }

        #endregion

        #region Tools

        public ObservableCollection<ToolViewModel> Tools
        {
            get
            {
                return Model.Tools;
            }
            set
            {
                if (Model.Tools != value)
                {
                    Model.Tools = value;
                    OnPropertyChanged(nameof(Tools));
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

        #region ActiveContent

        public object ActiveContent
        {
            get
            {
                return Model.ActiveContent;
            }
            set
            {
                if (Model.ActiveContent != value)
                {
                    Model.ActiveContent = value;

                    if (ActiveContent is DocumentViewModel)
                        ActiveDocument = ActiveContent as DocumentViewModel;

                    OnPropertyChanged(nameof(ActiveContent));
                }
            }
        }

        #endregion

        #region Explorer

        public ExplorerViewModel Explorer
        {
            get
            {
                return Model.Explorer;
            }
            set
            {
                if (Model.Explorer != value)
                {
                    Model.Explorer = value;
                    OnPropertyChanged(nameof(Explorer));
                }
            }
        }

        #endregion

        #region Console

        public ConsoleViewModel Console
        {
            get
            {
                return Model.Console;
            }
            set
            {
                if (Model.Console != value)
                {
                    Model.Console = value;
                    OnPropertyChanged(nameof(Console));
                }
            }
        }

        #endregion
    }
}
