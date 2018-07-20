/*
 * Mark Diedericks
 * 02/07/2018
 * Version 1.0.5
 * Theme Interface, Basic Members
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Excel_Macros_UI.Themes
{
    public interface ITheme
    {
        /// <summary>
        /// Basic Functions for Theme
        /// </summary>

        IList<Uri> UriList { get; }
        string Name { get; }

    }
}
