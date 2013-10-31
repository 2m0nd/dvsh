using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;

namespace DimondVSHelpers
{
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
        public Connect()
        {
        }

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;
            if (connectMode == ext_ConnectMode.ext_cm_UISetup)
            {
                object[] contextGUIDS = new object[] { };
                Commands2 commands = (Commands2)_applicationObject.Commands;
                string toolsMenuName = "Tools";

                //Place the command on the tools menu.
                //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
                Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

                //Find the Tools command bar on the MenuBar command bar:
                CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
                CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

                //This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
                //  just make sure you also update the QueryStatus/Exec method to include the new command names.
                try
                {
                    //Add a command to the Commands collection:
                    var showNotExistFilesCommand = commands.AddNamedCommand2(
                        _addInInstance,
                        "ShowNotExistCommand",
                        "ShowNotExistCommand",
                        "Executes the command for ShowNotExistCommand",
                        false,
                        Resources.ShowNotExistCommandImage,
                        ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                        (int)vsCommandStyle.vsCommandStylePictAndText,
                        vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for the command to the tools menu:
                    if ((showNotExistFilesCommand != null) && (toolsPopup != null))
                    {
                        showNotExistFilesCommand.AddControl(toolsPopup.CommandBar, 1);
                    }

                    //Add a command to the Commands collection:
                    var openBinaryCommand = commands.AddNamedCommand2(
                        _addInInstance,
                        "OpenBinaryCommand",
                        "OpenBinaryCommand",
                        "Executes the command for OpenBinaryCommand",
                        false,
                        Resources.OpenBinaryCommandImage,
                        ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                        (int)vsCommandStyle.vsCommandStylePictAndText,
                        vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for the command to the tools menu:
                    if ((openBinaryCommand != null) && (toolsPopup != null))
                    {
                        openBinaryCommand.AddControl(toolsPopup.CommandBar, 1);
                    }
                }
                catch (System.ArgumentException)
                {
                    //If we are here, then the exception is probably because a command with that name
                    //  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
                }
            }
        }

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param term='commandName'>The name of the command to determine state for.</param>
        /// <param term='neededText'>Text that is needed for the command.</param>
        /// <param term='status'>The state of the command in the user interface.</param>
        /// <param term='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (commandName == "DimondVSHelpers.Connect.ShowNotExistCommand")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
                if (commandName == "DimondVSHelpers.Connect.OpenBinaryCommand")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }

            }
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param term='commandName'>The name of the command to execute.</param>
        /// <param term='executeOption'>Describes how the command should be run.</param>
        /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param term='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName == "DimondVSHelpers.Connect.ShowNotExistCommand")
                {
                    handled = true;
                    ShowNotExistFiles();
                }
                if (commandName == "DimondVSHelpers.Connect.OpenBinaryCommand")
                {
                    handled = true;
                    OpenBinaryCommand();
                }
            }
        }

        void OpenBinaryCommand()
        {
            var pr = GetActiveProject(_applicationObject);

            var fullapath = pr.FullName;

            var dir = Path.GetDirectoryName(fullapath);

            var binaryFolder = Path.Combine(dir, "bin", pr.ConfigurationManager.ActiveConfiguration.ConfigurationName);

            System.Diagnostics.Process.Start("explorer", binaryFolder);

            return;
        }

        void ShowNotExistFiles()
        {
            var listNotExistedFiles = new List<String>();

            var solutionFiles = GetAllSolutionProjectItems();

            foreach (var projectItem in solutionFiles)
            {
                var item = (projectItem as ProjectItem);
                var fileNames = item.FileNames[0];
                var isExist = File.Exists(fileNames);
                if (!isExist && !fileNames.EndsWith("\\") && fileNames.Contains("\\"))
                {
                    listNotExistedFiles.Add(fileNames);
                }
            }


            // Find the output window.
            Window outputWindow = _applicationObject.Windows.Item(Constants.vsWindowKindOutput);
            // Show the window. (You might want to make sure outputWindow is not null here...)
            outputWindow.Visible = true;
            var outWin = (OutputWindow)outputWindow.Object;
            outWin.ActivePane.Activate();

            if (listNotExistedFiles.Any())
            {
                outWin.ActivePane.OutputString(string.Format("Count not found files :: {0}\n", listNotExistedFiles.Count));
                int k = 1;
                foreach (var notExistPath in listNotExistedFiles)
                {
                    outWin.ActivePane.OutputString(string.Format("{0}. {1}\n", k, notExistPath));
                    k++;
                }
            }
            else
            {
                outWin.ActivePane.OutputString(string.Format("All files exist in solution\n"));
            }
        }

        List<ProjectItem> GetAllSolutionProjectItems()
        {
            var projects = _applicationObject.Solution.Projects.GetEnumerator();
            var files = new List<ProjectItem>();
            while (projects.MoveNext())
            {
                var items = ((Project)projects.Current).ProjectItems.GetEnumerator();
                while (items.MoveNext())
                {
                    var item = (ProjectItem)items.Current;
                    //Recursion to get all ProjectItems
                    files.AddRange(GetFiles(item));
                }
            }
            return files;
        }

        List<ProjectItem> GetFiles(ProjectItem item)
        {
            var files = new List<ProjectItem>();
            //base case
            if (item.ProjectItems == null || item.ProjectItems.Count == 0)
            {
                files.Add(item);
                return files;
            }

            var items = item.ProjectItems.GetEnumerator();
            while (items.MoveNext())
            {
                var currentItem = (ProjectItem)items.Current;
                files.AddRange(GetFiles(currentItem));
            }

            return files;
        }

        static Project GetActiveProject(DTE2 dte)
        {
            Project activeProject = null;

            Array activeSolutionProjects = dte.ActiveSolutionProjects as Array;
            if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
            {
                activeProject = activeSolutionProjects.GetValue(0) as Project;
            }

            return activeProject;
        }

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
    }
}