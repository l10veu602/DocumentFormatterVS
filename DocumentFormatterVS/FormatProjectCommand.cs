//------------------------------------------------------------------------------
// <copyright file="FormatProjectCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace DocumentFormatterVS
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FormatProjectCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("48465474-d64b-4daa-9b7c-c60e0b7e3290");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly VSPackage1 package;

        private readonly DTE DTE;

        private readonly DocumentFormatter documentFormatter;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatProjectCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private FormatProjectCommand(VSPackage1 package, DTE DTE, DocumentFormatter documentFormatter)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            this.DTE = DTE;
            this.documentFormatter = documentFormatter;

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
        public static FormatProjectCommand Instance
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
        public static void Initialize(VSPackage1 package, DTE DTE, DocumentFormatter documentFormatter)
        {
            Instance = new FormatProjectCommand(package, DTE, documentFormatter);
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
            var activeSolutionProjects = DTE.ActiveSolutionProjects as object[];
            if (activeSolutionProjects == null || activeSolutionProjects.Length == 0)
            {
                return;
            }

            var selectedProject = activeSolutionProjects[0] as Project;
            if (selectedProject == null)
            {
                return;
            }

            foreach (var projectItem in selectedProject.ProjectItems)
            {
                FormatProjectItem(projectItem as ProjectItem);
            }
        }

        private void FormatProjectItem(ProjectItem projectItem)
        {
            if (projectItem == null)
            {
                return;
            }

            if (projectItem.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile && !projectItem.IsOpen)
            {
                // vcxproj.filter 등에 대한 예외 처리
                try
                {
                    projectItem.Open();
                }
                catch { }
            }

            if (projectItem.Document != null)
            {
                documentFormatter.FormatDocument(projectItem.Document);
            }

            foreach(var subProjectItem in projectItem.ProjectItems)
            {
                FormatProjectItem(subProjectItem as ProjectItem);
            }
        }
    }
}
