//------------------------------------------------------------------------------
// <copyright file="EnableCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace DocumentFormatterVS
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class EnableCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("48465474-d64b-4daa-9b7c-c60e0b7e3290");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly VSPackage1 package;

        /// <summary>
        /// Initializes a new instance of the <see cref="EnableCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private EnableCommand(VSPackage1 package)
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
                var menuItem = new OleMenuCommand(this.MenuItemCallback, null, this.BeforeQueryStatus, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private void BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuItem = sender as OleMenuCommand;
            menuItem.Text = package.IsFormattingEnabled ? "Formatting: Turn off" : "Formatting: Turn On";
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static EnableCommand Instance
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
        public static void Initialize(VSPackage1 package)
        {
            Instance = new EnableCommand(package);
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
            package.IsFormattingEnabled = !package.IsFormattingEnabled;
        }
    }
}
