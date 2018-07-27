using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Excel_Macros_UI.Utilities
{
    public static class FocusUtility
    {
        public static readonly DependencyProperty IsFocusedProperty = DependencyProperty.RegisterAttached("IsFocused", typeof(bool), typeof(FocusUtility), new UIPropertyMetadata(false, OnIsFocusedPropertyChanged));

        public static bool GetIsFocused(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsFocusedProperty);
        }

        public static void SetIsFocused(DependencyObject obj, bool value)
        {
            obj.SetValue(IsFocusedProperty, value);
        }

        private static void OnIsFocusedPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            UIElement uie = obj as UIElement;

            if ((bool)args.NewValue)
                uie.Focus();
        }
    }
}
