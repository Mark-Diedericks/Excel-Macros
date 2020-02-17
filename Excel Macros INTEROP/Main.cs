/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.15
 * The main hub of the interop library
 */

using Excel_Macros_INTEROP.Engine;
using Excel_Macros_INTEROP.Libraries;
using Excel_Macros_INTEROP.Macros;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Macros_INTEROP
{
    public class Main
    {
        #region Excel Application

        private static Excel.Application s_ExcelApplication;
        private static Dispatcher s_ExcelDispatcher;

        /// <summary>
        /// Gets the Excel application
        /// </summary>
        /// <returns></returns>
        public static Excel.Application GetApplication()
        {
            return s_ExcelApplication;
        }

        /// <summary>
        /// Gets the UI Dispatcher for the Excel application
        /// </summary>
        /// <returns></returns>
        public static Dispatcher GetApplicationDispatcher()
        {
            return s_ExcelDispatcher;
        }

        #endregion

        #region Initialization & Destruction

        //OnLoaded event, for all Forms and GUIs
        public delegate void OnLoadedEvent();
        public event OnLoadedEvent OnLoaded;

        //OnDestroyed event, for all Forms and GUIs
        public delegate void DestroyEvent();
        public event DestroyEvent OnDestroyed;

        //OnFocused event, for all Forms and GUIs
        public delegate void FocusEvent();
        public event FocusEvent OnFocused;

        //OnShown event, for all Forms and GUIs
        public delegate void ShowEvent();
        public event ShowEvent OnShown;

        //OnHidden event, for all Forms and GUIs
        public delegate void HideEvent();
        public event HideEvent OnHidden;

        //OnIOChanged event, for all Forms and GUIs
        public delegate void IOChangedEvent();
        public event IOChangedEvent OnIOChanged;

        //OnAssembliesChanged event, for all Forms and GUIs
        public delegate void AssembliesChangedEvent();
        public event AssembliesChangedEvent OnAssembliesChanged;

        //OnMacroCountChanged event, for all Forms and GUIs
        public delegate void MacroCountChangedEvent();
        public event MacroCountChangedEvent OnMacroCountChanged;

        //MacroRenamed event, for all Forms and GUIs
        public delegate void MacroRenameEvent(Guid id);
        public event MacroRenameEvent OnMacroRenamed;

        //Temporary path storage
        private string[] m_RibbonMacroPaths;

        //Macros
        private Dictionary<Guid, MacroDeclaration> m_Declarations;
        private Dictionary<Guid, IMacro> m_Macros;
        private HashSet<Guid> m_RibbonMacros;
        private Guid m_ActiveMacro;

        //User Included Assemblies
        private HashSet<AssemblyDeclaration> m_Assemblies;

        //IO Management
        private EngineIOManager m_EngineIOManager;

        //Instancing
        private static Main s_Instance;

        /// <summary>
        /// Public instantiation of Main
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <param name="dispatcher">Excel UI dispatcher</param>
        /// <param name="OnLoaded">Action to be fire when the task is completed</param>
        /// <param name="RibbonMacros">Serialized ribbon macro list</param>
        /// <param name="ActiveMacroRelativePath">Relative path of the last active macro</param>
        /// <param name="Libraries">Serialized HashSet of included assemblies</param>
        public static void Initialize(Excel.Application application, Dispatcher dispatcher, Action OnLoaded, string RibbonMacros, string ActiveMacroRelativePath, AssemblyDeclaration[] Libraries)
        {
            //Set local reference to excel application
            s_ExcelApplication = application;
            s_ExcelDispatcher = dispatcher;

            Dispatcher ui = Dispatcher.CurrentDispatcher;

            Task.Run(() =>
            {
                //Create Instance
                Main m = new Main();

                //Initialize Utilities and Managers
                EventManager.Instantiate();
                MessageManager.Instantiate();
                Utilities.Instantiate();

                //Initialize Execution Engine
                ExecutionEngine.Initialize();

                //Load saved macros
                Dictionary<MacroDeclaration, IMacro> macros = FileManager.LoadAllMacros(new string[] { FileManager.MacroDirectory });
                GetInstance().m_Declarations = new Dictionary<Guid, MacroDeclaration>();
                GetInstance().m_Macros = new Dictionary<Guid, IMacro>();

                for (int i = 0; i < macros.Count; i++)
                {
                    MacroDeclaration md = macros.Keys.ElementAt<MacroDeclaration>(i);
                    md.id = Guid.NewGuid();

                    IMacro im = macros[md];
                    im.SetID(md.id);

                    GetInstance().m_Declarations.Add(md.id, md);
                    GetInstance().m_Macros.Add(md.id, im);
                }

                GetInstance().OnMacroCountChanged?.Invoke();

                //Parse ribbon macros
                GetInstance().m_RibbonMacros = new HashSet<Guid>();
                GetInstance().m_RibbonMacroPaths = RibbonMacros.Split(';');

                //Get the active macro
                if (!String.IsNullOrEmpty(ActiveMacroRelativePath))
                    GetInstance().m_ActiveMacro = GetIDFromRelativePath(ActiveMacroRelativePath);
                else
                    GetInstance().m_ActiveMacro = GetInstance().m_Macros.Keys.FirstOrDefault<Guid>();

                //Get Assemblies
                if (Libraries != null)
                    GetInstance().m_Assemblies = new HashSet<AssemblyDeclaration>(Libraries);
                else
                    GetInstance().m_Assemblies = new HashSet<AssemblyDeclaration>();

                GetInstance().OnAssembliesChanged?.Invoke();

                ui.BeginInvoke(DispatcherPriority.Normal, new Action(() => {
                    OnLoaded?.Invoke();
                    GetInstance().OnLoaded?.Invoke();
                }));
            });
        }

        /// <summary>
        /// Fires OnDestroyed event
        /// </summary>
        public static void Destroy()
        {
            GetInstance().OnDestroyed?.Invoke();
        }

        /// <summary>
        /// Gets instance of Main
        /// </summary>
        /// <returns>Main instance</returns>
        public static Main GetInstance()
        {
            return s_Instance;
        }

        /// <summary>
        /// Private instantiation of Main
        /// </summary>
        private Main()
        {
            s_Instance = this;
        }

        #endregion

        #region Getters, Setters and Passthrough Functions

        /// <summary>
        /// Loads all ribbon macros from serialized list
        /// </summary>
        public static void LoadRibbonMacros()
        {
            GetInstance().m_RibbonMacros.Clear();

            foreach (string file in GetInstance().m_RibbonMacroPaths)
                AddRibbonMacro(GetIDFromRelativePath(file));
        }

        /// <summary>
        /// Gets the ID of the active macro
        /// </summary>
        /// <returns>ID of active macro</returns>
        public static Guid GetActiveMacro()
        {
            return s_Instance.m_ActiveMacro;
        }

        /// <summary>
        /// Sets the active macro
        /// </summary>
        /// <param name="macro">The macro's ID</param>
        public static void SetActiveMacro(Guid macro)
        {
            s_Instance.m_ActiveMacro = macro;
        }

        /// <summary>
        /// Sets the TextWriters for both output and error of the execution engines
        /// </summary>
        /// <param name="output">TextWriter for ouput stream</param>
        /// <param name="error">TextWriter for error stream</param>
        public static void SetIOSteams(TextWriter output, TextWriter error)
        {
            GetInstance().m_EngineIOManager = new EngineIOManager(output, error);
            GetInstance().OnIOChanged?.Invoke();
        }

        /// <summary>
        /// Gets the EngineIOManager instance
        /// </summary>
        /// <returns></returns>
        public static EngineIOManager GetEngineIOManager()
        {
            return GetInstance().m_EngineIOManager;
        }

        /// <summary>
        /// Checks if a macro is ribbon accessible 
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <returns></returns>
        public static bool IsRibbonMacro(Guid id)
        {
            return GetInstance().m_RibbonMacros.Contains(id);
        }

        /// <summary>
        /// Adds a macro to the ribbon
        /// </summary>
        /// <param name="id">The macro's id</param>
        public static void AddRibbonMacro(Guid id)
        {
            if (id == Guid.Empty || IsRibbonMacro(id))
                return;

            GetInstance().m_RibbonMacros.Add(id);

            MacroDeclaration md = GetInstance().m_Declarations[id];
            IMacro macro = GetInstance().m_Macros[id];

            EventManager.AddRibbonMacro(id, md.name, md.relativepath, () => ExecutionEngine.GetEngine().ExecuteMacro(macro.GetSource(), null, false));
        }

        /// <summary>
        /// Removes a macro from the ribbon
        /// </summary>
        /// <param name="id">The macro's id</param>
        public static void RemoveRibbonMacro(Guid id)
        {
            GetInstance().m_RibbonMacros.Remove(id);
            EventManager.RemoveRibbonMacro(id);
        }

        /// <summary>
        /// Renames a ribbon macro
        /// </summary>
        /// <param name="id">The macro's id</param>
        public static void RenameRibbonMacro(Guid id)
        {
            GetInstance().m_RibbonMacros.Add(id);

            MacroDeclaration md = GetInstance().m_Declarations[id];
            EventManager.RenameRibbonMacro(id, md.name, md.relativepath);
        }

        /// <summary>
        /// Fires Show and Focus events
        /// </summary>
        public static void FireShowFocusEvent()
        {
            FireShowEvent();
            FireFocusEvent();
        }

        /// <summary>
        /// Fires Show event
        /// </summary>
        public static void FireShowEvent()
        {
            GetInstance().OnShown?.Invoke();
        }

        /// <summary>
        /// Fires Focus event
        /// </summary>
        public static void FireFocusEvent()
        {
            GetInstance().OnFocused?.Invoke();
        }

        /// <summary>
        /// Fires Hide event
        /// </summary>
        public static void FireHideEvent()
        {
            GetInstance().OnHidden?.Invoke();
        }

        /// <summary>
        /// Sets Excel's interactivity state
        /// </summary>
        /// <param name="enabled">Whether or not Excel should be set as interactive</param>
        public static void SetExcelInteractive(bool enabled)
        {
            EventManager.ExcelSetInteractive(enabled);
        }

        /// <summary>
        /// Adds an assembly to the assembly list
        /// </summary>
        /// <param name="ad">AssemblyDeclaration of assembly</param>
        public static void AddAssembly(AssemblyDeclaration ad)
        {
            GetInstance().m_Assemblies.Add(ad);
            GetInstance().OnAssembliesChanged?.Invoke();
        }

        /// <summary>
        /// Removes an assembly from the assembly list
        /// </summary>
        /// <param name="ad">AssemblyDeclaration of the assembly</param>
        public static void RemoveAssembly(AssemblyDeclaration ad)
        {
            GetInstance().m_Assemblies.Remove(ad);
            GetInstance().OnAssembliesChanged?.Invoke();
        }

        /// <summary>
        /// Gets the list of Assemblies
        /// </summary>
        /// <returns>Gets list (HashSet) of assemblies</returns>
        public static HashSet<AssemblyDeclaration> GetAssemblies()
        {
            return GetInstance().m_Assemblies;
        }

        /// <summary>
        /// Gets an AssemblyDeclaration from its longname
        /// </summary>
        /// <param name="longname">An assembly's longname</param>
        /// <returns>The respective AssemblyDeclaration</returns>
        public static AssemblyDeclaration GetAssemblyByLongName(string longname)
        {
            foreach (AssemblyDeclaration ad in GetInstance().m_Assemblies)
                if (ad.filepath == longname)
                    return ad;

            return null;
        }

        /// <summary>
        /// Gets the list of macro declarations
        /// </summary>
        /// <returns>Dictionary of MacroDeclarations to their respective IDs</returns>
        public static Dictionary<Guid, MacroDeclaration> GetDeclarations()
        {
            return GetInstance().m_Declarations;
        }

        /// <summary>
        /// Gets the list of macro objects
        /// </summary>
        /// <returns>Dictionary of Macros and their respective IDs</returns>
        public static Dictionary<Guid, IMacro> GetMacros()
        {
            return GetInstance().m_Macros;
        }

        /// <summary>
        /// Gets a macro by its id
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <returns>Macro of the given id</returns>
        public static IMacro GetMacro(Guid id)
        {
            if (!GetInstance().m_Macros.ContainsKey(id))
                return null;

            return GetInstance().m_Macros[id];
        }

        /// <summary>
        /// Gets a MacroDeclaration from a macro's id
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <returns>MacroDeclaration of the given id</returns>
        public static MacroDeclaration GetDeclaration(Guid id)
        {
            if (!GetInstance().m_Declarations.ContainsKey(id))
                return null;

            return GetInstance().m_Declarations[id];
        }

        /// <summary>
        /// Gets the ID of a macro from it's relative path
        /// </summary>
        /// <param name="relativepath">The macro's relative path</param>
        /// <returns>Guid of the macro</returns>
        public static Guid GetIDFromRelativePath(string relativepath)
        {
            string path = relativepath.ToLower().Trim();

            foreach (MacroDeclaration macro in GetInstance().m_Declarations.Values)
                if (macro.relativepath.ToLower().Trim() == path)
                    return macro.id;

            return Guid.Empty;
        }

        /// <summary>
        /// Sets the MacroDeclaration associated with an ID
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <param name="declaration">The macro's MacroDeclaration</param>
        public static void SetDeclaration(Guid id, MacroDeclaration declaration)
        {
            if (!GetInstance().m_Declarations.ContainsKey(id))
                GetInstance().m_Declarations.Add(id, declaration);
            else
                GetInstance().m_Declarations[id] = declaration;
        }

        /// <summary>
        /// Adds a macro to the registry
        /// </summary>
        /// <param name="declaration">The macro's macro declaration</param>
        /// <param name="macro">The macro</param>
        /// <returns>The macro's assigned ID</returns>
        public static Guid AddMacro(MacroDeclaration declaration, IMacro macro)
        {
            Guid id = Guid.NewGuid();

            declaration.id = id;
            macro.SetID(id);

            GetInstance().m_Declarations.Add(id, declaration);
            GetInstance().m_Macros.Add(id, macro);
            GetInstance().OnMacroCountChanged?.Invoke();

            return id;
        }

        /// <summary>
        /// Removes a macro from the registry
        /// </summary>
        /// <param name="id">The macro's id</param>
        public static void RemoveMacro(Guid id)
        {
            GetInstance().m_Macros.Remove(id);
            GetInstance().OnMacroCountChanged?.Invoke();
        }

        /// <summary>
        /// Renames a macro
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <param name="newname">The macro's new name</param>
        public static void RenameMacro(Guid id, string newname)
        {
            if (!GetInstance().m_Macros.ContainsKey(id))
            {
                MessageManager.DisplayOkMessage("Could not find the macro: " + GetDeclaration(id).name, "Rename Macro Error");
                return;
            }

            IMacro macro = GetInstance().m_Macros[id];

            macro.Save();
            macro.Rename(newname);
            macro.Save();

            GetInstance().OnMacroRenamed?.Invoke(id);
        }

        /// <summary>
        /// Renames a folder
        /// </summary>
        /// <param name="olddir">The folder's current relative path</param>
        /// <param name="newdir">The folder's desired relative path</param>
        /// <returns>A list (HashSet) of ids of effected macros</returns>
        public static HashSet<Guid> RenameFolder(string olddir, string newdir)
        {
            HashSet<Guid> affectedMacros = new HashSet<Guid>();

            FileManager.RenameFolder(olddir, newdir);
            string relativepath = FileManager.CalculateRelativePath(FileManager.CalculateFullPath(olddir));

            foreach (Guid id in GetInstance().m_Declarations.Keys)
            {
                if (GetDeclaration(id).relativepath.ToLower().Trim().StartsWith(relativepath.ToLower().Trim()))
                {
                    affectedMacros.Add(id);
                    GetInstance().m_Declarations[id].relativepath = GetDeclaration(id).relativepath.Replace(relativepath, FileManager.CalculateRelativePath(FileManager.CalculateFullPath(newdir)));
                }
            }

            return affectedMacros;
        }

        /// <summary>
        /// Deletes a folder
        /// </summary>
        /// <param name="directory">The relative path of the folder</param>
        /// <param name="OnReturn">The Action, a bool representing the operations success, to be fired when the task is completed</param>
        public static void DeleteFolder(string directory, Action<bool> OnReturn)
        {
            HashSet<Guid> affectedMacros = new HashSet<Guid>();

            FileManager.DeleteFolder(directory, new Action<bool>((result) =>
            {
                if (!result)
                    OnReturn?.Invoke(false);

                string relativepath = FileManager.CalculateRelativePath(FileManager.CalculateFullPath(directory)).ToLower().Trim();

                HashSet<Guid> toremove = new HashSet<Guid>();
                foreach (Guid id in GetInstance().m_Declarations.Keys)
                    if (GetDeclaration(id).relativepath.ToLower().Trim().Contains(relativepath))
                        toremove.Add(id);

                foreach (Guid id in toremove)
                    GetInstance().m_Declarations.Remove(id);

                OnReturn?.Invoke(true);
            }));
        }

        #endregion
    }

}
