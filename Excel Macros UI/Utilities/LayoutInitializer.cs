using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Xceed.Wpf.AvalonDock.Layout;

namespace Excel_Macros_UI.Utilities
{
    public enum PaneLocation
    {
        Left = 0,
        Right = 1,
        Bottom = 2
    }

    internal class LayoutInitializer : ILayoutUpdateStrategy
    {
        private static string GetPaneLocationName(PaneLocation location)
        {
            switch(location)
            {
                case PaneLocation.Left:
                    return "LeftPane";
                case PaneLocation.Right:
                    return "RightPane";
                case PaneLocation.Bottom:
                    return "BottomPane";
                default:
                    return "RightPane";
            }
        }

        private static LayoutAnchorablePane CreateAnchorablePane(LayoutRoot layout, Orientation orientation, string location, bool start)
        {
            LayoutPanel parent = layout.Descendents().OfType<LayoutPanel>().First(x => x.Orientation == orientation);
            LayoutAnchorablePane pane = new LayoutAnchorablePane() { Name = location };

            if (start)
                parent.InsertChildAt(0, pane);
            else
                parent.Children.Add(pane);

            return pane;
        }

        public bool BeforeInsertAnchorable(LayoutRoot layout, LayoutAnchorable anchorableToShow, ILayoutContainer destinationContainer)
        {
            ToolViewModel tool = anchorableToShow.Content as ToolViewModel;
            if (tool != null)
            {
                PaneLocation prefferedLocation = tool.PreferredLocation;
                string location = GetPaneLocationName(prefferedLocation);
                LayoutAnchorablePane pane = layout.Descendents().OfType<LayoutAnchorablePane>().FirstOrDefault(x => x.Name == location);

                if(pane == null)
                {
                    switch(prefferedLocation)
                    {
                        case PaneLocation.Left:
                            pane = CreateAnchorablePane(layout, Orientation.Horizontal, location, true);
                            break;
                        case PaneLocation.Right:
                            pane = CreateAnchorablePane(layout, Orientation.Horizontal, location, false);
                            break;
                        case PaneLocation.Bottom:
                            pane = CreateAnchorablePane(layout, Orientation.Vertical, location, false);
                            break;
                        default:
                            pane = CreateAnchorablePane(layout, Orientation.Horizontal, location, false);
                            break;
                    }
                }

                pane.Children.Add(anchorableToShow);
                return true;
            }

            return false;
        }

        public void AfterInsertAnchorable(LayoutRoot layout, LayoutAnchorable anchorableShown)
        {
            
        }

        public bool BeforeInsertDocument(LayoutRoot layout, LayoutDocument anchorableToShow, ILayoutContainer destinationContainer)
        {
            return false;
        }

        public void AfterInsertDocument(LayoutRoot layout, LayoutDocument anchorableShown)
        {
            
        }

    }
}
