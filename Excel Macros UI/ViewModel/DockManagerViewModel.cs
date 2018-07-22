/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * Handles the view models of the primary view model
 */

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
        public ObservableCollection<DocumentViewModel> Documents { get; set; }
        public ObservableCollection<ToolViewModel> Tools { get; set; }

        public DockManagerViewModel(IEnumerable<DocumentViewModel> DocumentViewModels, IEnumerable<ToolViewModel> ToolViewModels)
        {
            Documents = new ObservableCollection<DocumentViewModel>();
            Tools = new ObservableCollection<ToolViewModel>();

            foreach(DocumentViewModel document in DocumentViewModels)
            {
                document.PropertyChanged += Document_PropertyChanged;

                if (!document.IsClosed)
                    Documents.Add(document);
            }

            foreach (ToolViewModel tool in ToolViewModels)
            {
                tool.PropertyChanged += Tool_PropertyChanged;

                if (!tool.IsClosed)
                    Tools.Add(tool);
            }
        }

        private void Tool_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ToolViewModel tool = (ToolViewModel)sender;

            if (e.PropertyName == nameof(ToolViewModel.IsClosed))
            {
                if (!tool.IsClosed)
                    Tools.Add(tool);
                else
                    Tools.Remove(tool);
            }
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
    }
}
