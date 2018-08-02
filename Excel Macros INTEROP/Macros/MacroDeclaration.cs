/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.3
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
    /// <summary>
    /// Data strcture containing info on a macro, serializable data structure for saving
    /// </summary>
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

    /// <summary>
    /// Enum identifying macro type
    /// </summary>
    [Serializable]
    public enum MacroType
    {
        PYTHON = 0,
    }

    /// <summary>
    /// Converter to serialize MacroDeclaration instances
    /// </summary>
    public class MacroDeclarationConverter : TypeConverter
    {
        /// <summary>
        /// Interface method, ensures that the source can be deserialized into an MacroDeclaration
        /// </summary>
        /// <param name="context"></param>
        /// <param name="sourceType"></param>
        /// <returns>Bool identifying if it can be converted</returns>
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string);
        }

        /// <summary>
        /// Interface method, deserializes a string into an MacroDeclaration instance
        /// </summary>
        /// <param name="context"></param>
        /// <param name="culture"></param>
        /// <param name="value"></param>
        /// <returns>The MacroDeclaration that has been deserialized</returns>
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

        /// <summary>
        /// Interface method, serializes MacroDeclaration as string
        /// </summary>
        /// <param name="context"></param>
        /// <param name="culture"></param>
        /// <param name="value"></param>
        /// <param name="destinationType"></param>
        /// <returns>The string of the serialized MacroDeclaration</returns>
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
