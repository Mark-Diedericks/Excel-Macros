/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
 * Assembly declaration data structure
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Libraries
{
    [TypeConverter(typeof(AssemblyDeclarationConverter))]
    [SettingsSerializeAs(SettingsSerializeAs.String)]
    public class AssemblyDeclaration
    {
        public string displayname;
        public string filepath;

        public AssemblyDeclaration(string dn, string ln)
        {
            displayname = dn;
            filepath = ln;
        }
    }

    public class AssemblyDeclarationConverter : TypeConverter
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
                AssemblyDeclaration assembly = new AssemblyDeclaration(parts.Length > 0 ? parts[0] : "", parts.Length > 2 ? parts[2] : "");
                return assembly;
            }

            return base.ConvertFrom(context, culture, value);
        }

        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                AssemblyDeclaration assembly = value as AssemblyDeclaration;
                return string.Format("{0},{1}", assembly.displayname, assembly.filepath);
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }

    }
}
