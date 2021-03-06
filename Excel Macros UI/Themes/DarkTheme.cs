﻿/*
 * Mark Diedericks
 * 20/07/2018
 * Version 1.0.5
 * Dark Theme For UI
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Themes
{
    public sealed class DarkTheme : ITheme
    {
        /// <summary>
        /// Instantiate a new DarkTheme
        /// </summary>
        public DarkTheme()
        {
            UriList = new List<Uri>
            {
                new Uri("pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseDark.xaml"),
                new Uri("pack://application:,,,/Excel Macros UI;component/Themes/ExcelAccent.xaml"),
                new Uri("pack://application:,,,/AvalonDock.Themes.VS2012;component/DarkTheme.xaml"),
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
                return "Dark";
            }
        }
    }
}
