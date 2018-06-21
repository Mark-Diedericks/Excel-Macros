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
            tools.Add(new FileExplorerViewModel() { Title = "File Explorer", CanClose = true, ContentId = "FileExplorer" });
            tools.Add(new ToolboxViewModel() { Title = "Toolbox", CanClose = true, ContentId = "Toolbox" });
            tools.Add(new ConsoleViewModel() { Title = "Console", CanClose = true, ContentId = "Console" });

            List<DocumentViewModel> documents = new List<DocumentViewModel>();
            documents.Add(new TextualEditorViewModel() { Title = "Textual Editor", CanFloat = true, ContentId = "TestTextDoc1" });
            documents.Add(new VisualEditorViewModel() { Title = "Visual Editor", CanFloat = true, ContentId = "TestVisDoc1" });

            DockManagerViewModel = new DockManagerViewModel(documents, tools);

            //DockManagerViewModel = new DockManagerViewModel();
        }
    }
}
