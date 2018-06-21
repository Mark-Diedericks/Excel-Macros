/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
 * Manages all system file related tasks
 */

using Excel_Macros_INTEROP.Macros;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Excel_Macros_INTEROP
{
    public class FileManager
    {

        public static readonly string PYTHON_FILE_EXT = ".ipy";
        public static readonly string PYTHON_FILTER = "Python Macro | *" + PYTHON_FILE_EXT;

        public static readonly string BLOCKLY_FILE_EXT = ".xml";
        public static readonly string BLOCKLY_FILTER = "Node Macro | *" + BLOCKLY_FILE_EXT;
        
        public static readonly string MACRO_FILTER = "Macro | *" + BLOCKLY_FILE_EXT + ", " + PYTHON_FILE_EXT;

        #region DIRECTORIES

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = System.Reflection.Assembly.GetAssembly(typeof(FileManager)).CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public static string MacroDirectory
        {
            get
            {
                return Path.GetFullPath(AssemblyDirectory + "/Macros/");
            }
        }

        #endregion

        #region MACRO_LOADING

        public static List<string> GetFiles(string directory)
        {
            if (!Directory.Exists(directory))
                return new List<string>();

            List<string> result = new List<string>();
            string[] files = Directory.GetFiles(directory);
            string[] dirs = Directory.GetDirectories(directory);

            foreach (string file in files)
                result.Add(file);

            foreach (string dir in dirs)
                result.AddRange(GetFiles(dir));
            return result;
        }

        public static List<MacroDeclaration> IdentifyAllMacros(string[] directories)
        {
            List<MacroDeclaration> declarations = new List<MacroDeclaration>();

            foreach (string directory in directories)
            {
                List<string> files = GetFiles(directory);

                foreach (string file in files)
                {
                    if (Path.GetExtension(file).ToLower().Trim() == PYTHON_FILE_EXT)
                    {
                        string relativepath = CalculateRelativePath(file);
                        string fullpath = CalculateFullPath(relativepath);

                        FileInfo fi = new FileInfo(fullpath);
                        if (!fi.Directory.Exists)
                            fi.Directory.Create();

                        declarations.Add(new MacroDeclaration(MacroType.PYTHON, Path.GetFileName(fullpath), relativepath));
                    }
                }
            }
            return declarations;
        }

        public static Dictionary<MacroDeclaration, IMacro> LoadAllMacros(string[] directories)
        {
            Dictionary<MacroDeclaration, IMacro> macros = new Dictionary<MacroDeclaration, IMacro>();
            List<MacroDeclaration> declarations = IdentifyAllMacros(directories);

            foreach (MacroDeclaration md in declarations)
            {
                IMacro macro = LoadMacro(md.type, md.relativepath);

                if (macro != null)
                    macros.Add(md, macro);
            }

            return macros;
        }

        #endregion

        #region GENERIC_FILE_UTIL

        public static void SaveMacro(Guid id, string source)
        {
            if (Main.GetDeclaration(id) == null)
                return;

            try
            {
                string fullpath = CalculateFullPath(Main.GetDeclaration(id).relativepath);

                FileInfo fi = new FileInfo(fullpath);
                if (!fi.Directory.Exists)
                    fi.Directory.Create();

                File.WriteAllText(fullpath, source);
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not save macro: \"" + Main.GetDeclaration(id).name + "\". \n\n" + e.Message, "Saving Error");
            }
        }

        public static IMacro LoadMacro(MacroType type, string relativepath)
        {
            switch (type)
            {
                case MacroType.PYTHON: return LoadPythonScript(relativepath);
                case MacroType.BLOCKLY: return LoadBlocklyScript(relativepath);
            }

            return null;
        }

        public static void ExportMacro(Guid id, string source)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = Main.GetDeclaration(id).name;
            sfd.Filter = Main.GetDeclaration(id).type == MacroType.PYTHON ? PYTHON_FILTER : PYTHON_FILTER;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //if (Main.GetMacroManager().Visibility == System.Windows.Visibility.Visible)
                //    Main.GetMacroManager().TryFocus();

                Main.FireShowFocusEvent();

                try
                {
                    File.WriteAllText(sfd.FileName, source);
                }
                catch (Exception e)
                {
                    DisplayOkMessage("Could not export macro: \"" + Main.GetDeclaration(id).name + "\". \n\n" + e.Message, "Saving Error");
                }
            }

            //if (Main.GetMacroManager().Visibility == System.Windows.Visibility.Visible)
            //    Main.GetMacroManager().TryFocus();

            Main.FireShowFocusEvent();
        }

        public static void ImportMacro(string relativedir, Action<Guid> OnReturn)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = PYTHON_FILTER;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //if (Main.GetMacroManager().Visibility == System.Windows.Visibility.Visible)
                //    Main.GetMacroManager().TryFocus();

                Main.FireShowFocusEvent();

                bool pyext = Path.GetExtension(ofd.FileName).ToLower().Trim() == PYTHON_FILE_EXT.ToLower().Trim();
                MacroType macroType = pyext ? MacroType.PYTHON : MacroType.PYTHON;

                string newpath = CalculateFullPath(relativedir + ofd.SafeFileName);

                System.Diagnostics.Debug.WriteLine(newpath);

                string relativepath = CalculateRelativePath(newpath);
                string fullpath = CalculateFullPath(relativepath);

                FileInfo fi = new FileInfo(fullpath);
                if (!fi.Directory.Exists)
                    fi.Directory.Create();

                if (File.Exists(fullpath))
                {
                    DisplayYesNoMessage("This file already exists, would you like to replace it?", "File Overwrite", new Action<bool>((result) =>
                    {
                        if (!result)
                            OnReturn?.Invoke(Guid.Empty);

                        File.Copy(ofd.FileName, fullpath, true);

                        MacroDeclaration declaration = new MacroDeclaration(macroType, ofd.SafeFileName, relativepath);
                        IMacro macro = LoadMacro(macroType, relativepath);

                        OnReturn?.Invoke(Main.AddMacro(declaration, macro));
                    }));
                }
            }

            //if (Main.GetMacroManager().Visibility == System.Windows.Visibility.Visible)
            //    Main.GetMacroManager().TryFocus();

            Main.FireShowFocusEvent();

            OnReturn?.Invoke(Guid.Empty);
        }

        public static bool RenameMacro(Guid id, string name)
        {
            try
            {
                string newpath = Main.GetDeclaration(id).relativepath.Replace(Main.GetDeclaration(id).name, name);

                MacroDeclaration declaration = new MacroDeclaration(Main.GetDeclaration(id).type, name, newpath);
                declaration.id = id;

                File.Move(CalculateFullPath(Main.GetDeclaration(id).relativepath), CalculateFullPath(declaration.relativepath));
                Main.SetDeclaration(id, declaration);

                return true;
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not rename the macro file: " + Main.GetDeclaration(id).name + "\n" + e.Message, "Renaming Error");
            }

            return false;
        }

        public static Guid CreateMacro(MacroType type, string relativepath)
        {
            try
            {
                MacroDeclaration declaration = new MacroDeclaration(type, Path.GetFileName(relativepath), relativepath);

                File.CreateText(CalculateFullPath(relativepath)).Close();

                IMacro macro = LoadMacro(type, relativepath);
                macro.CreateBlankMacro();

                return Main.AddMacro(declaration, macro);
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not create the macro file: " + Path.GetFileName(relativepath) + "\n" + e.Message, "Creation error");
            }

            return Guid.Empty;
        }

        public static bool CreateFolder(string relativepath)
        {
            try
            {
                if (!Directory.Exists(CalculateFullPath(relativepath)))
                    Directory.CreateDirectory(CalculateFullPath(relativepath));

                return true;
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not create the folder: " + relativepath.Replace('\\', '/').Replace("//", "/") + "\n" + e.Message, "Creation error");
            }

            return false;
        }

        public static void DeleteFolder(string relativepath, Action<bool> OnReturn)
        {
            DisplayYesNoMessage("'" + relativepath + "' Will be deleted permanently.", "Macro Deletion", new Action<bool>((result) => {
                if (result)
                {
                    try
                    {
                        Directory.Delete(CalculateFullPath(relativepath), true);
                        OnReturn?.Invoke(true);
                    }
                    catch (Exception e)
                    {
                        DisplayOkMessage("Could not delete the folder: " + relativepath + "\n" + e.Message, "Creation error");
                    }
                }
            }));


            OnReturn?.Invoke(false);
        }

        public static bool RenameFolder(string oldpath, string newpath)
        {
            try
            {
                Directory.Move(CalculateFullPath(oldpath), CalculateFullPath(newpath));
                return true;
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not rename the folder: " + oldpath + "\n" + e.Message, "Creation error");
            }

            return false;
        }

        public static void DeleteMacro(Guid id, Action<bool> OnReturn)
        {
            DisplayYesNoMessage("'" + Main.GetDeclaration(id).name + "' Will be deleted permanently.", "Macro Deletion", new Action<bool>(result =>
            { 
                if (result)
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Deleting...");
                        File.Delete(CalculateFullPath(Main.GetDeclaration(id).relativepath));
                        Main.RemoveMacro(id);

                        OnReturn?.Invoke(true);
                    }
                    catch (Exception e)
                    {
                        DisplayOkMessage("Could not delete macro: \"" + Main.GetDeclaration(id).name + "\". \n\n" + e.Message, "Deletion Error");
                    }
                }
                }));

            OnReturn?.Invoke(false);
        }

        public static string CalculateRelativePath(string fullpath)
        {
            return fullpath.Remove(0, MacroDirectory.Length);
        }

        public static string CalculateFullPath(string relativepath)
        {
            return Path.GetFullPath(MacroDirectory + relativepath);
        }

        #endregion

        #region PYTHON_FILE_UTIL

        public static readonly string[] PYTHON_LIB_PATH = { "./PythonModules/ctypes/", "./PythonModules/distutils/", "./PythonModules/email/", "./PythonModules/encodings/", "./PythonModules/ensurepip/", "./PythonModules/importlib/", "./PythonModules/json/", "./PythonModules/lib2to3/", "./PythonModules/logging/", "./PythonModules/multiprocessing/", "./PythonModules/pydoc_data/", "./PythonModules/site-packages", "./PythonModules/sqlite3/", "./PythonModules/unitest/", "./PythonModules/wsgrief/", "./PythonModules/xml/" };

        private static TextualMacro LoadPythonScript(string relativepath)
        {
            try
            {
                string fullpath = CalculateFullPath(relativepath);

                FileInfo fi = new FileInfo(fullpath);
                if (!fi.Exists)
                    return null;

                string source = File.ReadAllText(fullpath.Trim());
                return new TextualMacro(source);
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not open macro: \"" + relativepath + "\". \n\n" + e.Message, "Loading Error");
            }

            return null;
        }

        private static VisualMacro LoadBlocklyScript(string relativepath)
        {
            try
            {
                string fullpath = CalculateFullPath(relativepath);

                FileInfo fi = new FileInfo(fullpath);
                if (!fi.Exists)
                    return null;

                string source = File.ReadAllText(fullpath.Trim());
                return new VisualMacro(source);
            }
            catch (Exception e)
            {
                DisplayOkMessage("Could not open macro: \"" + relativepath + "\". \n\n" + e.Message, "Loading Error");
            }

            return null;
        }

        #endregion

        private static void DisplayOkMessage(string message, string caption)
        {
            /*if (Main.GetMacroManager() == null)
                MessageBox.Show(message, caption, MessageBoxButtons.OK);
            else
                Main.GetMacroManager().DisplayOkMessage(message, caption);*/

            MessageManager.DisplayOkMessage(message, caption);
        }

        private static void DisplayYesNoMessage(string message, string caption, Action<bool> OnReturn)
        {
            /*if (Main.GetMacroManager() == null)
                return MessageBox.Show(message, caption, MessageBoxButtons.YesNo) == DialogResult.Yes;
            else
                return await Main.GetMacroManager().DisplayYesNoMessage(message, caption);*/

            MessageManager.DisplayYesNoMessage(message, caption, OnReturn);
        }

    }
}
