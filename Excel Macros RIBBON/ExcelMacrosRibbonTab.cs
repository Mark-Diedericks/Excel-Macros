/*
 * Mark Diedericks
 * 08/06/2018
 * Version 1.0.0
 * The ribbon tab UI for the AddIn
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Macros_RIBBON
{
    public partial class ExcelMacrosRibbonTab
    {
        private struct RibbonMacroData
        {
            public string name;
            public string path;
        }

        private static ExcelMacrosRibbonTab s_Instance = null;
        private Dictionary<Guid, RibbonButton> m_Buttons;
        private Dictionary<Guid, RibbonMacroData> m_Macros;

        public delegate void MacroRibbonEvent();
        public static event MacroRibbonEvent MacroRibbonLoadEvent;
        public event MacroRibbonEvent MacroEditorClickEvent;
        public event MacroRibbonEvent NewTextualClickEvent;
        public event MacroRibbonEvent NewVisualClickEvent;
        public event MacroRibbonEvent OpenMacroClickEvent;

        /// <summary>
        /// Returns the instance of the Excel Tab of the AddIn
        /// </summary>
        /// <returns>Instance of the object</returns>
        public static ExcelMacrosRibbonTab GetInstance()
        {
            return s_Instance;
        }

        #region Loading

        private void ExcelMacrosRibbonTab_Load(object sender, RibbonUIEventArgs e)
        {
            s_Instance = this;

            SetUIState(false); //Disabled UI whilst we load
            groupMacros.Label = "Loading...";

            m_Buttons = new Dictionary<Guid, RibbonButton>();
            m_Macros = new Dictionary<Guid, RibbonMacroData>();

            MacroRibbonLoadEvent?.Invoke();
        }

        public void MainUILoaded()
        {
            SetUIState(true);
            groupMacros.Label = "Macros";
        }

        #endregion

        #region UI Updates/States

        /// <summary>
        /// Sets the enabled state of the ribbon tab's UI elements
        /// </summary>
        /// <param name="state">Enabled or Disabled state (True/False)</param>
        private void SetUIState(bool state)
        {
            menuExecuteMacro.Enabled = state && menuExecuteMacro.Items.Count > 0;

            btnMacroEditor.Enabled = state;
            btnNewTextual.Enabled = state;
            btnNewVisual.Enabled = state;
            btnOpenMacro.Enabled = state;
        }

        private void UpdateRibbon()
        {
            if (menuExecuteMacro.Items.Count <= 0)
                menuExecuteMacro.Enabled = false;
        }

        #endregion
        
        #region Manage Executable Macros

        /// <summary>
        /// Returns a serialized list of the ribbon accessible macros
        /// </summary>
        /// <returns></returns>
        public string GetRibbonMacros()
        {
            StringBuilder sb = new StringBuilder();

            foreach(Guid id in m_Macros.Keys)
            {
                RibbonMacroData macro = m_Macros[id];
                sb.Append(macro.path);
                sb.Append(';');
            }

            return sb.ToString();
        }

        /// <summary>
        /// Removes a Macro's respective button UI control from the drop down menu
        /// </summary>
        /// <param name="id">ID of the Macro</param>
        public void RemoveMacro(Guid id)
        {
            if (m_Buttons == null)
                return;

            if (!m_Buttons.ContainsKey(id))
                return;

            RibbonButton button = m_Buttons[id];

            if (menuExecuteMacro.Items.Contains(button))
                menuExecuteMacro.Items.Remove(button);

            m_Buttons.Remove(id);
            m_Macros.Remove(id);

            UpdateRibbon();
        }

        /// <summary>
        /// Add a Macro's respective button UI control to the drop down menu
        /// </summary>
        /// <param name="id">ID of the macro</param>
        /// <param name="macroName">Name of the macro</param>
        /// <param name="macroPath">Relative path of the macro</param>
        /// <param name="macroClickEvent">Event to be called when macro is clicked</param>
        public void AddMacro(Guid id, string macroName, string macroPath, Action macroClickEvent)
        {
            if (m_Buttons == null)
                return;

            RibbonButton button = this.Factory.CreateRibbonButton();

            button.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            button.ShowImage = false;
            button.ShowLabel = true;

            button.Name = macroName;
            button.Label = macroName;
            button.ScreenTip = macroPath;

            button.Click += delegate (object sender, RibbonControlEventArgs args) { macroClickEvent.Invoke(); };

            m_Buttons.Add(id, button);
            m_Macros.Add(id, new RibbonMacroData() { name = macroName, path = macroPath});
            menuExecuteMacro.Items.Add(button);

            UpdateRibbon();
        }

        /// <summary>
        /// Rename a the button UI control of the respective macro which was renamed
        /// </summary>
        /// <param name="id">ID of the macro</param>
        /// <param name="macroName">Name of the macro</param>
        /// <param name="macroPath">Relative path of the macro</param>
        public void RenameMacro(Guid id, string macroName, string macroPath)
        {
            if (m_Buttons == null)
                return;

            if (!m_Buttons.ContainsKey(id))
                return;

            RibbonButton button = m_Buttons[id];

            if (!menuExecuteMacro.Items.Contains(button))
                menuExecuteMacro.Items.Add(button);

            int index = menuExecuteMacro.Items.IndexOf(button);

            ((RibbonButton)menuExecuteMacro.Items[index]).Label = macroName;
            ((RibbonButton)menuExecuteMacro.Items[index]).Name = macroName;
            ((RibbonButton)menuExecuteMacro.Items[index]).ScreenTip = macroPath;

            m_Macros[id] = new RibbonMacroData() { name = macroName, path = macroPath };

            UpdateRibbon();
        }

        #endregion

        #region Event Callbacks
        
        private void btnMacroEditor_Click(object sender, RibbonControlEventArgs e)
        {
            MacroEditorClickEvent?.Invoke();
        }

        private void btnNewTextual_Click(object sender, RibbonControlEventArgs e)
        {
            NewTextualClickEvent?.Invoke();
        }

        private void btnNewVisual_Click(object sender, RibbonControlEventArgs e)
        {
            NewVisualClickEvent?.Invoke();
        }

        private void btnOpenMacro_Click(object sender, RibbonControlEventArgs e)
        {
            OpenMacroClickEvent?.Invoke();
        }

        #endregion
    }
}
