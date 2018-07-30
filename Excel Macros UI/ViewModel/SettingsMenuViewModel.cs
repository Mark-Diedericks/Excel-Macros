/*
 * Mark Diedericks
 * 30/06/2018
 * Version 1.0.1
 * Settings menu view model
 */

using Excel_Macros_UI.Model;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Macros_UI.Utilities.SyntaxStyleLoader;

namespace Excel_Macros_UI.ViewModel
{
    public class SettingsMenuViewModel : Base.ViewModel
    {
        private string PreviousTheme;
        private static SettingsMenuViewModel s_Instance;

        public SettingsMenuViewModel()
        {
            s_Instance = this;
            Model = new SettingsMenuModel();
            MainWindow.ThemeChanged += MainWindow_ThemeChanged;
            PreviousTheme = MainWindow.GetInstance().ActiveTheme.Name;
            LoadColors();
        }

        private void MainWindow_ThemeChanged()
        {
            SaveSyntaxStyle(PreviousTheme == "Dark");
            PreviousTheme = MainWindow.GetInstance().ActiveTheme.Name;
            LoadColors();
        }

        private void LoadColors()
        {
            LoadColorValues();
            string[] values = GetValues();

            FunctionColor = values[(int)SyntaxStyleColor.FUNCTION];
            DigitColor = values[(int)SyntaxStyleColor.DIGIT];
            CommentColor = values[(int)SyntaxStyleColor.COMMENT];
            StringColor = values[(int)SyntaxStyleColor.STRING];
            PairColor = values[(int)SyntaxStyleColor.PAIR];
            ClassColor = values[(int)SyntaxStyleColor.CLASS];
            StatementColor = values[(int)SyntaxStyleColor.STATEMENT];
            BooleanColor = values[(int)SyntaxStyleColor.BOOLEAN];
        }

        private void SetColors()
        {
            string[] values = new string[8];

            values[(int)SyntaxStyleColor.FUNCTION] = FunctionColor;
            values[(int)SyntaxStyleColor.DIGIT] = DigitColor;
            values[(int)SyntaxStyleColor.COMMENT] = CommentColor;
            values[(int)SyntaxStyleColor.STRING] = StringColor;
            values[(int)SyntaxStyleColor.PAIR] = PairColor;
            values[(int)SyntaxStyleColor.CLASS] = ClassColor;
            values[(int)SyntaxStyleColor.STATEMENT] = StatementColor;
            values[(int)SyntaxStyleColor.BOOLEAN] = BooleanColor;

            SetSyntaxStyle(values);
        }

        #region Model

        private SettingsMenuModel m_Model;
        public SettingsMenuModel Model
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

        #region FunctionColor

        public string FunctionColor
        {
            get
            {
                return Model.FunctionColor;
            }

            set
            {
                if (Model.FunctionColor != value)
                {
                    Model.FunctionColor = value;
                    OnPropertyChanged(nameof(FunctionColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region DigitColor
        
        public string DigitColor
        {
            get
            {
                return Model.DigitColor;
            }

            set
            {
                if (Model.DigitColor != value)
                {
                    Model.DigitColor = value;
                    OnPropertyChanged(nameof(DigitColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region CommentColor
        
        public string CommentColor
        {
            get
            {
                return Model.CommentColor;
            }

            set
            {
                if (Model.CommentColor != value)
                {
                    Model.CommentColor = value;
                    OnPropertyChanged(nameof(CommentColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region StringColor
        
        public string StringColor
        {
            get
            {
                return Model.StringColor;
            }

            set
            {
                if (Model.StringColor != value)
                {
                    Model.StringColor = value;
                    OnPropertyChanged(nameof(StringColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region PairColor
        
        public string PairColor
        {
            get
            {
                return Model.PairColor;
            }

            set
            {
                if (Model.PairColor != value)
                {
                    Model.PairColor = value;
                    OnPropertyChanged(nameof(PairColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region ClassColor
        
        public string ClassColor
        {
            get
            {
                return Model.ClassColor;
            }

            set
            {
                if (Model.ClassColor != value)
                {
                    Model.ClassColor = value;
                    OnPropertyChanged(nameof(ClassColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region StatementColor
        
        public string StatementColor
        {
            get
            {
                return Model.StatementColor;
            }

            set
            {
                if (Model.StatementColor != value)
                {
                    Model.StatementColor = value;
                    OnPropertyChanged(nameof(StatementColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region BooleanColor
        
        public string BooleanColor
        {
            get
            {
                return Model.BooleanColor;
            }

            set
            {
                if (Model.BooleanColor != value)
                {
                    Model.BooleanColor = value;
                    OnPropertyChanged(nameof(BooleanColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region RibbonItems

        public ObservableCollection<DisplayableTreeViewItem> RibbonItems
        {
            get
            {
                return Model.RibbonItems;
            }
            set
            {
                if (Model.RibbonItems != value)
                {
                    Model.RibbonItems = value;
                    OnPropertyChanged(nameof(RibbonItems));
                }
            }
        }

        #endregion

        #region LabelVisible

        public bool LabelVisible
        {
            get
            {
                return Model.LabelVisible;
            }
            set
            {
                if(Model.LabelVisible != value)
                {
                    Model.LabelVisible = value;
                    OnPropertyChanged(nameof(LabelVisible));
                }
            }
        }

        #endregion

        #region LightTheme

        public bool LightTheme
        {
            get
            {
                return Model.LightTheme;
            }
            set
            {
                if (Model.LightTheme != value)
                {
                    Model.LightTheme = value;
                    OnPropertyChanged(nameof(LightTheme));
                    OnPropertyChanged(nameof(DarkTheme));
                }
            }
        }

        #endregion

        #region DarkTheme

        public bool DarkTheme
        {
            get
            {
                return Model.DarkTheme;
            }
            set
            {
                if (Model.DarkTheme != value)
                {
                    Model.DarkTheme = value;
                    OnPropertyChanged(nameof(LightTheme));
                    OnPropertyChanged(nameof(DarkTheme));
                }
            }
        }

        #endregion

        #region StyleActive

        public bool StyleActive
        {
            get
            {
                return Model.StyleActive;
            }
            set
            {
                if (Model.StyleActive != value)
                {
                    Model.StyleActive = value;
                    OnPropertyChanged(nameof(StyleActive));
                    OnPropertyChanged(nameof(LibraryActive));
                    OnPropertyChanged(nameof(RibbonActive));
                }
            }
        }

        #endregion

        #region LibraryActive

        public bool LibraryActive
        {
            get
            {
                return Model.LibraryActive;
            }
            set
            {
                if (Model.LibraryActive != value)
                {
                    Model.LibraryActive = value;
                    OnPropertyChanged(nameof(StyleActive));
                    OnPropertyChanged(nameof(LibraryActive));
                    OnPropertyChanged(nameof(RibbonActive));
                }
            }
        }

        #endregion

        #region RibbonActive

        public bool RibbonActive
        {
            get
            {
                return Model.RibbonActive;
            }
            set
            {
                if (Model.RibbonActive != value)
                {
                    Model.RibbonActive = value;
                    OnPropertyChanged(nameof(StyleActive));
                    OnPropertyChanged(nameof(LibraryActive));
                    OnPropertyChanged(nameof(RibbonActive));
                }
            }
        }

        #endregion
    }
}
