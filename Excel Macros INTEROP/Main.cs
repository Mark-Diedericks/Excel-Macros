/*
 * Mark Diedericks
 * 26/07/2018
 * Version 1.0.12
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

        public static Excel.Application GetApplication()
        {
            return s_ExcelApplication;
        }

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

        //OnMacroCountChanged event, for all Forms and GUIs
        public delegate void MacroCountChangedEvent();
        public event MacroCountChangedEvent OnMacroCountChanged;

        //MacroRenamed event, for all Forms and GUIs
        public delegate void MacroRenameEvent(Guid id);
        public event MacroRenameEvent OnMacroRenamed;

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

        public static void Initialize(Excel.Application application, Dispatcher dispatcher, Action OnLoaded, string RibbonMacros, string ActiveMacroRelativePath)
        {
            //Set local reference to excel application
            s_ExcelApplication = application;
            s_ExcelDispatcher = dispatcher;

            new Action(() =>
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

                string[] paths = RibbonMacros.Split(';');
                foreach(string file in paths)
                {
                    foreach (Guid key in GetInstance().m_Declarations.Keys)
                    {
                        string path = GetInstance().m_Declarations[key].relativepath;

                        if (path.Trim().ToLower() == file.Trim().ToLower())
                            AddRibbonMacro(key);
                    }
                }

                //Get the active macro
                if (!String.IsNullOrEmpty(ActiveMacroRelativePath))
                    GetInstance().m_ActiveMacro = GetIDFromRelativePath(ActiveMacroRelativePath);
                else
                    GetInstance().m_ActiveMacro = GetInstance().m_Macros.Keys.FirstOrDefault<Guid>();

                //Get Assemblies
                //if (Properties.Settings.Default.IncludedAssemblies != null)
                //    GetInstance().m_Assemblies = new HashSet<AssemblyDeclaration>(Properties.Settings.Default.IncludedAssemblies);
                //else
                GetInstance().m_Assemblies = new HashSet<AssemblyDeclaration>();
                
            }).BeginInvoke(new AsyncCallback((result) => {
                OnLoaded?.Invoke();
                GetInstance().OnLoaded?.Invoke();
                }), null);
        }

        public static void Destroy()
        {
            GetInstance().OnDestroyed?.Invoke();
        }

        public static Main GetInstance()
        {
            return s_Instance;
        }

        public Main()
        {
            s_Instance = this;
        }

        #endregion

        #region Getters, Setters and Passthrough Functions

        public static Guid GetActiveMacro()
        {
            return s_Instance.m_ActiveMacro;
        }

        public static void SetActiveMacro(Guid macro)
        {
            s_Instance.m_ActiveMacro = macro;
        }

        public static void SetIOSteams(TextWriter output, TextWriter error)
        {
            GetInstance().m_EngineIOManager = new EngineIOManager(output, error);
            GetInstance().OnIOChanged?.Invoke();
        }

        public static EngineIOManager GetEngineIOManager()
        {
            return GetInstance().m_EngineIOManager;
        }

        public static bool IsRibbonMacro(Guid id)
        {
            return GetInstance().m_RibbonMacros.Contains(id);
        }

        public static void AddRibbonMacro(Guid id)
        {
            GetInstance().m_RibbonMacros.Add(id);

            MacroDeclaration md = GetInstance().m_Declarations[id];
            IMacro macro = GetInstance().m_Macros[id];

            EventManager.AddRibbonMacro(id, md.name, md.relativepath, () => ExecutionEngine.GetReleaseEngine().ExecuteMacro(macro.GetSource(), null, false));
        }

        public static void RemoveRibbonMacro(Guid id)
        {
            GetInstance().m_RibbonMacros.Remove(id);
            EventManager.RemoveRibbonMacro(id);
        }

        public static void RenameRibbonMacro(Guid id)
        {
            GetInstance().m_RibbonMacros.Add(id);

            MacroDeclaration md = GetInstance().m_Declarations[id];
            EventManager.RenameRibbonMacro(id, md.name, md.relativepath);
        }

        public static void FireShowFocusEvent()
        {
            FireShowEvent();
            FireFocusEvent();
        }

        public static void FireShowEvent()
        {
            GetInstance().OnShown?.Invoke();
        }

        public static void FireFocusEvent()
        {
            GetInstance().OnFocused?.Invoke();
        }

        public static void FireHideEvent()
        {
            GetInstance().OnHidden?.Invoke();
        }

        public static void SetExcelInteractive(bool enabled)
        {
            EventManager.ExcelSetInteractive(enabled);
        }

        public static void AddAssembly(AssemblyDeclaration ad)
        {
            GetInstance().m_Assemblies.Add(ad);
        }

        public static void RemoveAssembly(AssemblyDeclaration ad)
        {
            GetInstance().m_Assemblies.Remove(ad);
        }

        public static HashSet<AssemblyDeclaration> GetAssemblies()
        {
            return GetInstance().m_Assemblies;
        }

        public static AssemblyDeclaration GetAssemblyByLongName(string longname)
        {
            foreach (AssemblyDeclaration ad in GetInstance().m_Assemblies)
                if (ad.filepath == longname)
                    return ad;

            return null;
        }

        public static Dictionary<Guid, MacroDeclaration> GetDeclarations()
        {
            return GetInstance().m_Declarations;
        }

        public static Dictionary<Guid, IMacro> GetMacros()
        {
            return GetInstance().m_Macros;
        }

        public static IMacro GetMacro(Guid id)
        {
            if (!GetInstance().m_Macros.ContainsKey(id))
                return null;

            return GetInstance().m_Macros[id];
        }

        public static MacroDeclaration GetDeclaration(Guid id)
        {
            if (!GetInstance().m_Declarations.ContainsKey(id))
                return null;

            return GetInstance().m_Declarations[id];
        }

        public static Guid GetIDFromRelativePath(string relativepath)
        {
            string path = relativepath.ToLower().Trim();

            foreach (MacroDeclaration macro in GetInstance().m_Declarations.Values)
                if (macro.relativepath.ToLower().Trim() == path)
                    return macro.id;

            return Guid.Empty;
        }

        public static void SetDeclaration(Guid id, MacroDeclaration declaration)
        {
            if (!GetInstance().m_Declarations.ContainsKey(id))
                GetInstance().m_Declarations.Add(id, declaration);
            else
                GetInstance().m_Declarations[id] = declaration;
        }

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

        public static void RemoveMacro(Guid id)
        {
            GetInstance().m_Macros.Remove(id);
            GetInstance().OnMacroCountChanged?.Invoke();
        }

        public static Guid GetGuidFromRelativePath(string relativepath)
        {
            string ltrp = relativepath.ToLower().Trim();

            foreach (Guid id in GetInstance().m_Declarations.Keys)
                if (GetDeclaration(id).relativepath.ToLower().Trim().Equals(ltrp))
                    return id;

            return Guid.Empty;
        }

        public static Guid RenameMacro(Guid id, string newname)
        {
            if (!GetInstance().m_Macros.ContainsKey(id))
            {
                MessageManager.DisplayOkMessage("Could not find the macro: " + GetDeclaration(id).name, "Rename Macro Error");
                return id;
            }

            IMacro macro = GetInstance().m_Macros[id];

            macro.Save();
            macro.Rename(newname);
            macro.Save();

            GetInstance().OnMacroRenamed?.Invoke(id);

            return id;
        }

        public static void RenameFolder(string olddir, string newdir)
        {
            HashSet<Guid> affectedMacros = new HashSet<Guid>();

            FileManager.RenameFolder(olddir, newdir);
            string relativepath = FileManager.CalculateRelativePath(FileManager.CalculateFullPath(olddir));

            foreach (Guid id in GetInstance().m_Declarations.Keys)
                if (GetDeclaration(id).relativepath.ToLower().Trim().StartsWith(relativepath.ToLower().Trim()))
                    GetInstance().m_Declarations[id].relativepath = GetDeclaration(id).relativepath.Replace(relativepath, FileManager.CalculateRelativePath(FileManager.CalculateFullPath(newdir)));
        }

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
