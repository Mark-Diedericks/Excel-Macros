/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Primary view model for handling main window's views
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class PrimaryViewModel
    {
        public DockManagerViewModel DockManagerViewModel { get; private set; }

        public PrimaryViewModel()
        {
            List<ToolViewModel> tools = new List<ToolViewModel>();
            tools.Add(new FileExplorerViewModel() { Title = "File Explorer", CanClose = true });
            
            List<DocumentViewModel> documents = new List<DocumentViewModel>();
            documents.Add(new TextualEditorViewModel() { Title = "Textual Editor", CanFloat = true });
            documents.Add(new VisualEditorViewModel() { Title = "Visual Editor", CanFloat = true });

            DockManagerViewModel = new DockManagerViewModel(documents, tools);
        }
    }
}
