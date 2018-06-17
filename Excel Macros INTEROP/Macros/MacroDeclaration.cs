/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.1
 * Macro declaration data structure
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.ComponentModel;

namespace Excel_Macros_INTEROP.Macros
{
    [TypeConverter(typeof(MacroDeclarationConverter))]
    [SettingsSerializeAs(SettingsSerializeAs.String)]
    public class MacroDeclaration
    {
        public MacroType type;
        public string name;
        public string relativepath;
        public Guid id = Guid.Empty;

        public MacroDeclaration(MacroType t, string n, string p)
        {
            type = t;
            name = n;
            relativepath = p;
        }
    }

    [Serializable]
    public enum MacroType
    {
        PYTHON = 0,
        BLOCKLY = 1
    }

    public class MacroDeclarationConverter : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
        {
            if (value is string)
            {
                string[] parts = ((string)value).Split(new char[] { ',' });
                MacroDeclaration macro = new MacroDeclaration((MacroType)Convert.ToInt32(parts[0]), parts.Length > 1 ? parts[1] : "", parts.Length > 2 ? parts[2] : "");
                return macro;
            }

            return base.ConvertFrom(context, culture, value);
        }

        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                MacroDeclaration macro = value as MacroDeclaration;
                return string.Format("{0},{1},{2},{3}", ((int)macro.type).ToString(), macro.name, macro.relativepath);
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }

    }
}
