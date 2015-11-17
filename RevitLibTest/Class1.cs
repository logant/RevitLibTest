using System;
using System.Collections.Generic;
using System.Windows.Interop;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using adWin = Autodesk.Windows;


namespace RevitLibTest 
{
    /// <summary>
    /// This command displays the scale of a revit document compared to meters, such as 1.0 for meters or 0.3048 for feet.
    /// This command also uses a function in a stand alone library to get the document scale and return it.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RevitScaleCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                double scale = LINE.RevitLib2016.RevitScaleToMeters(commandData.Application.ActiveUIDocument.Document);
                TaskDialog.Show("Revit Scale", "Revit Scale Compared to Meters: " + scale.ToString());
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// This ExternalApplication acts as an example in how we have typically implemented UI integration of commands
    /// </summary>
    public class TraditionalRevitScaleApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Get the path to the command
                string path = typeof(RevitScaleApp).Assembly.Location;

                // Verify if the tab exists, create it if ncessary
                adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

                string tabName = "HKS";
                string panelName = "Tools";

                adWin.RibbonTab tab = null;
                foreach (adWin.RibbonTab t in ribbon.Tabs)
                {
                    if (t.Id == tabName)
                    {
                        tab = t;
                        break;
                    }
                }
                if (tab == null)
                    application.CreateRibbonTab(tabName);

                // Verify if the panel exists
                List<RibbonPanel> panels = application.GetRibbonPanels(tabName);
                RibbonPanel panel = null;
                foreach (RibbonPanel rp in panels)
                {
                    if (rp.Name == panelName)
                    {
                        panel = rp;
                        break;
                    }
                }
                if (panel == null)
                    panel = application.CreateRibbonPanel(tabName, panelName);

                // Construct the PushButtonData
                PushButtonData scalePBD = new PushButtonData(
                    "Revit Scale", "Revit\nScale", path, "RevitLibTest.RevitScaleCmd")
                {
                    LargeImage = Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.icon.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()),
                    ToolTip = "Get a number value compare the document scale to 1m.",
                };

                // Add the butotn to the panel
                panel.AddItem(scalePBD);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// This ExternalApplication acts as an example in how to implement UI integration of commands by using an external library
    /// </summary>
    public class RevitScaleApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Get the path to the command
                string path = typeof(RevitScaleApp).Assembly.Location;
                
                // Construct the PushButtonData
                PushButtonData scalePBD = new PushButtonData(
                    "Revit Scale", "Revit\nScale", path, "RevitLibTest.RevitScaleCmd")
                {
                    LargeImage = Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.icon.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()),
                    ToolTip = "Get a number value compare the document scale to 1m.",
                };

                // Add the scalePBD to a list.  There could be an override option for the AddToRibbon method of RevitLib 
                // manage creating a single versus a list of buttons.
                List<PushButtonData> buttons = new List<PushButtonData> { scalePBD };

                // Create the button(s)
                LINE.RevitLib2016.AddToRibbon(application, "HKS", "Tools", buttons);

                return Result.Succeeded;
            }
            catch 
            {
                return Result.Failed;
            }
        }
    }

    
}
