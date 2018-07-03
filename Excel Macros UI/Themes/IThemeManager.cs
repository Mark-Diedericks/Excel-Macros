﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Themes
{
    public interface IThemeManager
    {

        ObservableCollection<ITheme> Themes { get; }
        ITheme ActiveTheme { get; }

        bool AddTheme(ITheme theme);
        bool SetTheme(string name);

    }
}