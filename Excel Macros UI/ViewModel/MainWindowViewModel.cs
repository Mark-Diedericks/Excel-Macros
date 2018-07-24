/*
 * Mark Diedericks
 * 24/07/2018
 * Version 1.0.4
 * Primary view model for handling main window's views
 */

using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel.Base;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class MainWindowViewModel
    {
        public DockManagerViewModel DockManagerViewModel { get; private set; }

        public MainWindowViewModel()
        {
            List<DocumentViewModel> documents = new List<DocumentViewModel>();
            documents.Add(new TextualEditorViewModel() { Model = new TextualEditorModel() { Source = new TextDocument("ActiveWorksheet.Cells(1,1).Value = \"Hello\""), Title = "Textual Editor", ContentId = "TestTextDoc1", IsClosed = false },  CanFloat = true });
            documents.Add(new TextualEditorViewModel() { Model = new TextualEditorModel() { Source = new TextDocument("ActiveWorksheet.Cells(1,1).Value = \"Heyyy\""), Title = "Textual Editor", ContentId = "TestTextDoc2", IsClosed = false }, CanFloat = true });
            documents.Add(new VisualEditorViewModel() { Model = new VisualEditorModel() { Title = "Visual Editor", ContentId = "TestVisDoc1" }, CanFloat = true });

            DockManagerViewModel = new DockManagerViewModel(documents);
        }
    }
}
