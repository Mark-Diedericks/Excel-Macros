using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Utilities
{
    public class SyntaxStyleLoader
    {
        public enum SyntaxStyleColor
        {
            DIGIT = 0,
            COMMENT = 1,
            STRING = 2,
            PAIR = 3,
            CLASS = 4,
            STATEMENT = 5,
            FUNCTION = 6,
            BOOLEAN = 7
        }

        //Style Changed Event
        public delegate void StyleChangeEvent(Stream style);
        public static event StyleChangeEvent OnStyleChanged;

        //constants
        private const string DIGIT = "#COLOR_DIGIT";        //#202020
        private const string COMMENT = "#COLOR_COMMENT";      //#57a64a
        private const string STRING = "#COLOR_STRING";       //#ff22ff
        private const string PAIR = "#COLOR_PAIR";         //#569cd6
        private const string CLASS = "#COLOR_CLASS";        //#4ec9b0
        private const string STATEMENT = "#COLOR_STATEMENT";    //#70b0e0
        private const string FUNCTION = "#COLOR_FUNCTION";     //#404040
        private const string BOOLEAN = "#COLOR_BOOLEAN";      //#569cd6

        //Values
        private static string[] s_ColorValues;

        public static Stream GetStyleStream()
        {
            string style = Properties.Resources.IronPython;
            s_ColorValues = ParseSyntaxStyleString(Properties.Settings.Default.SyntaxStyle);

            style = style.Replace(DIGIT, s_ColorValues[(int)SyntaxStyleColor.DIGIT]);
            style = style.Replace(COMMENT, s_ColorValues[(int)SyntaxStyleColor.COMMENT]);
            style = style.Replace(STRING, s_ColorValues[(int)SyntaxStyleColor.STRING]);
            style = style.Replace(PAIR, s_ColorValues[(int)SyntaxStyleColor.PAIR]);
            style = style.Replace(CLASS, s_ColorValues[(int)SyntaxStyleColor.CLASS]);
            style = style.Replace(STATEMENT, s_ColorValues[(int)SyntaxStyleColor.STATEMENT]);
            style = style.Replace(FUNCTION, s_ColorValues[(int)SyntaxStyleColor.FUNCTION]);
            style = style.Replace(BOOLEAN, s_ColorValues[(int)SyntaxStyleColor.BOOLEAN]);

            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);

            writer.Write(style);
            writer.Flush();
            stream.Position = 0;

            return stream;
        }

        private static string CreateSyntaxStyleString(string[] values)
        {
            StringBuilder sb = new StringBuilder();

            foreach (string val in values)
                sb.Append(val + ';');

            return sb.ToString();
        }

        private static string[] ParseSyntaxStyleString(string value)
        {
            return value.Split(';');
        }

        private static void UpdateSyntaxStyle()
        {
            OnStyleChanged?.Invoke(GetStyleStream());
        }

        public static void SetSyntaxStyle(string[] values)
        {
            if (values.Length != 8)
                return;

            s_ColorValues = values;

            UpdateSyntaxStyle();
        }

        public static void SetSyntaxColor(SyntaxStyleColor color, string value)
        {
            s_ColorValues[(int)color] = value;
        }
    }
}
