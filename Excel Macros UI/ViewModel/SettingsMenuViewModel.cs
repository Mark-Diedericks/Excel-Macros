using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Macros_UI.Utilities.SyntaxStyleLoader;

namespace Excel_Macros_UI.ViewModel
{
    public class SettingsMenuViewModel : Base.ViewModel
    {
        private string PreviousTheme;

        public SettingsMenuViewModel()
        {
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

        #region FunctionColor

        private string m_FunctionColor;
        public string FunctionColor
        {
            get
            {
                return m_FunctionColor;
            }

            set
            {
                if (m_FunctionColor != value)
                {
                    m_FunctionColor = value;
                    OnPropertyChanged(nameof(FunctionColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region DigitColor

        private string m_DigitColor;
        public string DigitColor
        {
            get
            {
                return m_DigitColor;
            }

            set
            {
                if (m_DigitColor != value)
                {
                    m_DigitColor = value;
                    OnPropertyChanged(nameof(DigitColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region CommentColor

        private string m_CommentColor;
        public string CommentColor
        {
            get
            {
                return m_CommentColor;
            }

            set
            {
                if (m_CommentColor != value)
                {
                    m_CommentColor = value;
                    OnPropertyChanged(nameof(CommentColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region StringColor

        private string m_StringColor;
        public string StringColor
        {
            get
            {
                return m_StringColor;
            }

            set
            {
                if (m_StringColor != value)
                {
                    m_StringColor = value;
                    OnPropertyChanged(nameof(StringColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region PairColor

        private string m_PairColor;
        public string PairColor
        {
            get
            {
                return m_PairColor;
            }

            set
            {
                if (m_PairColor != value)
                {
                    m_PairColor = value;
                    OnPropertyChanged(nameof(PairColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region ClassColor

        private string m_ClassColor;
        public string ClassColor
        {
            get
            {
                return m_ClassColor;
            }

            set
            {
                if (m_ClassColor != value)
                {
                    m_ClassColor = value;
                    OnPropertyChanged(nameof(ClassColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region StatementColor

        private string m_StatementColor;
        public string StatementColor
        {
            get
            {
                return m_StatementColor;
            }

            set
            {
                if (m_StatementColor != value)
                {
                    m_StatementColor = value;
                    OnPropertyChanged(nameof(StatementColor));
                    SetColors();
                }
            }
        }

        #endregion

        #region BooleanColor

        private string m_BooleanColor;
        public string BooleanColor
        {
            get
            {
                return m_BooleanColor;
            }

            set
            {
                if (m_BooleanColor != value)
                {
                    m_BooleanColor = value;
                    OnPropertyChanged(nameof(BooleanColor));
                    SetColors();
                }
            }
        }

        #endregion

    }
}
