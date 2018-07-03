/*
 * Mark Diedericks
 * 02/07/2018
 * Version 1.0.5
 * Default Theme (Light Theme)
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Themes
{
    public sealed class DefaultTheme : ITheme
    {

        public DefaultTheme()
        {
            UriList = new List<Uri>
            {
                new Uri("pack://application:,,,/Excel Macros UI;component/Themes/LightTheme.xaml"),
                new Uri("pack://application:,,,/AvalonDock.Themes.VS2012;component/LightTheme.xaml")
            };
        }

        /// <summary>
        /// Inherited Memebers
        /// </summary>

        public IList<Uri> UriList { get; internal set; }
        public string Name
        {
            get
            {
                return "Default";
            }
        }

    }
}
