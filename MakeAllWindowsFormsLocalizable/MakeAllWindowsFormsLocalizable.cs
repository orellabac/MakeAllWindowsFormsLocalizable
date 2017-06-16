//------------------------------------------------------------------------------
// <copyright file="NewWebMapKendoWidget.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.ComponentModel;

namespace NewWebMapKendoWidget
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class MakeAllWindowsFormsLocalizable
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("af56546d-576a-4759-b8df-db2071fa72bc");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="MakeAllWindowsFormsLocalizable"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private MakeAllWindowsFormsLocalizable(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static MakeAllWindowsFormsLocalizable Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new MakeAllWindowsFormsLocalizable(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {

            var solution = ServiceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
            if (solution != null)
            {
                IntPtr hierarchyPointer, selectionContainerPointer;
                Object selectedObject = null;
                IVsMultiItemSelect multiItemSelect;
                uint projectItemId;

                IVsMonitorSelection monitorSelection =
                    (IVsMonitorSelection)Package.GetGlobalService(
                        typeof(SVsShellMonitorSelection));

                monitorSelection.GetCurrentSelection(out hierarchyPointer,
                    out projectItemId,
                    out multiItemSelect,
                    out selectionContainerPointer);

                IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                    hierarchyPointer,
                    typeof(IVsHierarchy)) as IVsHierarchy;

                if (selectedHierarchy == null)
                {
                    VsShellUtilities.ShowMessageBox(ServiceProvider, "Could not find current project", "Error",
                        OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK,OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                Project selectedProject = GetDTEProject(selectedHierarchy);
                
                if (selectedProject == null) return;

               // EnvDTE80.DTE2 m_dte2 = (EnvDTE80.DTE2)Package.GetGlobalService(typeof(EnvDTE.DTE));


                foreach (ProjectItem item in selectedProject.ProjectItems)
                {
                    try
                    {
                        var window = item.Open(EnvDTE.Constants.vsext_vk_Designer);
                        //objWindow = m_objDTE.ActiveDocument.ActiveWindow

                        // Get the designer host
                        var objIDesignerHost = window.Object as IDesignerHost;
                        if (objIDesignerHost != null)
                        {
                            //Get the container
                            var objIContainer = objIDesignerHost.Container;

                            //Iterate the components in a linear way (not hierarchically)
                            foreach (var objIComponent in objIContainer.Components)
                            {
                                if (objIComponent is Form)
                                {

                                    var colPropertyDescriptorCollection = TypeDescriptor.GetProperties(objIComponent);
                                    var objPropertyDescriptor = colPropertyDescriptorCollection["Localizable"];
                                    if (objPropertyDescriptor != null)
                                    {
                                        objPropertyDescriptor.SetValue(objIComponent, true);
                                    }
                                }
                                else if (objIComponent is UserControl)
                                {
                                    var colPropertyDescriptorCollection = TypeDescriptor.GetProperties(objIComponent);
                                    var objPropertyDescriptor = colPropertyDescriptorCollection["Localizable"];
                                    if (objPropertyDescriptor != null)
                                    {
                                        objPropertyDescriptor.SetValue(objIComponent, true);
                                    }
                                }


                            }
                        }
                    }
                    catch
                    {

                    }

                }

                /*string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                Directory.CreateDirectory(tempDirectory);
                var outFileName = "filename.cs";
                var completeFileName = Path.Combine(tempDirectory, outFileName);
                File.WriteAllText(completeFileName, "Hola");
                selectedProject.ProjectItems.AddFromFileCopy(completeFileName);*/
            }

            /*var form = new frmMain();
            if (form.ShowDialog() == DialogResult.OK)
            {
             
            }*/

        
        }



        /// <summary>
        /// Gets a list of projects in a solution.
        /// </summary>
        /// <param name="solution">The soloution to get the project list for.</param>
        /// <returns>List of projects.</returns>
        private IEnumerable<EnvDTE.Project> GetProjects(IVsSolution solution)
        {
            foreach (IVsHierarchy h in GetProjectsInSolution(solution, __VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION))
            {
                EnvDTE.Project project = GetDTEProject(h);
                if (project != null) yield return project;
            }
        }

        /// <summary>
        /// Gets a list of projects in a solution.
        /// </summary>
        /// <param name="solution">The soloution to get the project list for.</param>
        /// <param name="flags">Flags to specify which projects to get.</param>
        /// <returns></returns>
        private IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution, __VSENUMPROJFLAGS flags)
        {
            if (solution == null) yield break;

            IEnumHierarchies enumHierarchies;
            Guid guid = Guid.Empty;
            solution.GetProjectEnum((uint)flags, ref guid, out enumHierarchies);

            if (enumHierarchies == null) yield break;

            IVsHierarchy[] hierarchy = new IVsHierarchy[1];
            uint fetched;
            while (enumHierarchies.Next(1, hierarchy, out fetched) == VSConstants.S_OK && fetched == 1)
            {
                if (hierarchy.Length > 0 && hierarchy[0] != null)
                    yield return hierarchy[0];
            }
        }

        /// <summary>
        /// Gets a project from a hierarchy.
        /// </summary>
        /// <param name="hierarchy"></param>
        /// <returns></returns>
        public static EnvDTE.Project GetDTEProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null) throw new ArgumentNullException("hierarchy");

            object obj;
            hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ExtObject, out obj);
            return obj as EnvDTE.Project;
        }
    }
}
