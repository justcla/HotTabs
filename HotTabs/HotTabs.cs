//------------------------------------------------------------------------------
// <copyright file="HotTabs.cs" company="Justin Clareburt">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Linq;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.Shell;
using System.Diagnostics;

namespace HotTabs
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class HotTabs
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3b6bccd3-4c5a-4bcb-8ae4-8bf54dc3d642");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static HotTabs Instance
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
            Instance = new HotTabs(package);
        }
        
        /// <summary>
                 /// Initializes a new instance of the <see cref="HotTabs"/> class.
                 /// Adds our command handlers for menu (commands must exist in the command table file)
                 /// </summary>
                 /// <param name="package">Owner package, not null.</param>
        private HotTabs(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                int[] commandIDs = new int[]
                {
                    PkgCmdIDList.cmdidToolsOptions,
                    PkgCmdIDList.cmdidGoToPinnedTab1,
                    PkgCmdIDList.cmdidGoToPinnedTab2,
                    PkgCmdIDList.cmdidGoToPinnedTab3,
                    PkgCmdIDList.cmdidGoToPinnedTab4,
                    PkgCmdIDList.cmdidGoToPinnedTab5,
                    PkgCmdIDList.cmdidGoToPinnedTab6,
                    PkgCmdIDList.cmdidGoToPinnedTab7,
                    PkgCmdIDList.cmdidGoToPinnedTab8,
                    PkgCmdIDList.cmdidGoToPinnedTab9,
                    PkgCmdIDList.cmdidGoToPinnedTab10,
                    PkgCmdIDList.cmdidGoToUnpinnedTab1,
                    PkgCmdIDList.cmdidGoToUnpinnedTab2,
                    PkgCmdIDList.cmdidGoToUnpinnedTab3,
                    PkgCmdIDList.cmdidGoToUnpinnedTab4,
                    PkgCmdIDList.cmdidGoToUnpinnedTab5,
                    PkgCmdIDList.cmdidGoToUnpinnedTab6,
                    PkgCmdIDList.cmdidGoToUnpinnedTab7,
                    PkgCmdIDList.cmdidGoToUnpinnedTab8,
                    PkgCmdIDList.cmdidGoToUnpinnedTab9,
                    PkgCmdIDList.cmdidGoToUnpinnedTab10,
                };

                foreach (int id in commandIDs)
                {
                    CommandID menuCommandID = new CommandID(CommandSet, id);
                    OleMenuCommand menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                    menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
                    commandService.AddCommand(menuItem);
                }
            }
        }

        void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand command = (OleMenuCommand)sender;

            switch ((uint)command.CommandID.ID)
            {
                case PkgCmdIDList.cmdidToolsOptions:
                    // This one should always be enabled.
                    command.Visible = true;
                    command.Enabled = true;
                    break;

                case PkgCmdIDList.cmdidGoToPinnedTab1:
                case PkgCmdIDList.cmdidGoToPinnedTab2:
                case PkgCmdIDList.cmdidGoToPinnedTab3:
                case PkgCmdIDList.cmdidGoToPinnedTab4:
                case PkgCmdIDList.cmdidGoToPinnedTab5:
                case PkgCmdIDList.cmdidGoToPinnedTab6:
                case PkgCmdIDList.cmdidGoToPinnedTab7:
                case PkgCmdIDList.cmdidGoToPinnedTab8:
                case PkgCmdIDList.cmdidGoToPinnedTab9:
                case PkgCmdIDList.cmdidGoToPinnedTab10:
                    {
                        int desiredTab = command.CommandID.ID - PkgCmdIDList.cmdidGoToPinnedTab1;
                        Debug.Assert(desiredTab >= 0);

                        command.Visible = false;
                        command.Enabled = false;

                        View activeView = ViewManager.Instance.ActiveView;
                        if (null != activeView && DocumentGroup.IsTabbedDocument(activeView))
                        {
                            DocumentGroup group = activeView.Parent as DocumentGroup;
                            if (null != group && null != GetPinnedView(group, desiredTab))
                            {
                                command.Visible = true;
                                command.Enabled = true;
                            }
                        }
                        break;
                    }

                case PkgCmdIDList.cmdidGoToUnpinnedTab1:
                case PkgCmdIDList.cmdidGoToUnpinnedTab2:
                case PkgCmdIDList.cmdidGoToUnpinnedTab3:
                case PkgCmdIDList.cmdidGoToUnpinnedTab4:
                case PkgCmdIDList.cmdidGoToUnpinnedTab5:
                case PkgCmdIDList.cmdidGoToUnpinnedTab6:
                case PkgCmdIDList.cmdidGoToUnpinnedTab7:
                case PkgCmdIDList.cmdidGoToUnpinnedTab8:
                case PkgCmdIDList.cmdidGoToUnpinnedTab9:
                case PkgCmdIDList.cmdidGoToUnpinnedTab10:
                    {
                        int desiredTab = command.CommandID.ID - PkgCmdIDList.cmdidGoToUnpinnedTab1;
                        Debug.Assert(desiredTab >= 0);

                        command.Visible = false;
                        command.Enabled = false;

                        View activeView = ViewManager.Instance.ActiveView;
                        if (null != activeView && DocumentGroup.IsTabbedDocument(activeView))
                        {
                            DocumentGroup group = activeView.Parent as DocumentGroup;
                            if (null != group && null != GetUnpinnedView(group, desiredTab))
                            {
                                command.Visible = true;
                                command.Enabled = true;
                            }
                        }
                        break;
                    }
            }
        }

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            OleMenuCommand command = (OleMenuCommand)sender;

            switch (command.CommandID.ID)
            {
                //case PkgCmdIDList.cmdidToolsOptions:
                //    ShowMruPropertyPage();
                //    break;

                case PkgCmdIDList.cmdidGoToPinnedTab1:
                case PkgCmdIDList.cmdidGoToPinnedTab2:
                case PkgCmdIDList.cmdidGoToPinnedTab3:
                case PkgCmdIDList.cmdidGoToPinnedTab4:
                case PkgCmdIDList.cmdidGoToPinnedTab5:
                case PkgCmdIDList.cmdidGoToPinnedTab6:
                case PkgCmdIDList.cmdidGoToPinnedTab7:
                case PkgCmdIDList.cmdidGoToPinnedTab8:
                case PkgCmdIDList.cmdidGoToPinnedTab9:
                case PkgCmdIDList.cmdidGoToPinnedTab10:
                    {
                        int desiredTab = command.CommandID.ID - PkgCmdIDList.cmdidGoToPinnedTab1;
                        Debug.Assert(desiredTab >= 0);

                        View activeView = ViewManager.Instance.ActiveView;
                        if (null != activeView && DocumentGroup.IsTabbedDocument(activeView))
                        {
                            DocumentGroup group = activeView.Parent as DocumentGroup;
                            if (null != group)
                            {
                                View newView = GetPinnedView(group, desiredTab);
                                if (null != newView)
                                {
                                    newView.IsSelected = true;
                                }
                            }
                        }
                        break;
                    }

                case PkgCmdIDList.cmdidGoToUnpinnedTab1:
                case PkgCmdIDList.cmdidGoToUnpinnedTab2:
                case PkgCmdIDList.cmdidGoToUnpinnedTab3:
                case PkgCmdIDList.cmdidGoToUnpinnedTab4:
                case PkgCmdIDList.cmdidGoToUnpinnedTab5:
                case PkgCmdIDList.cmdidGoToUnpinnedTab6:
                case PkgCmdIDList.cmdidGoToUnpinnedTab7:
                case PkgCmdIDList.cmdidGoToUnpinnedTab8:
                case PkgCmdIDList.cmdidGoToUnpinnedTab9:
                case PkgCmdIDList.cmdidGoToUnpinnedTab10:
                    {
                        int desiredTab = command.CommandID.ID - PkgCmdIDList.cmdidGoToUnpinnedTab1;
                        Debug.Assert(desiredTab >= 0);

                        View activeView = ViewManager.Instance.ActiveView;
                        if (null != activeView && DocumentGroup.IsTabbedDocument(activeView))
                        {
                            DocumentGroup group = activeView.Parent as DocumentGroup;
                            if (null != group)
                            {
                                View newView = GetUnpinnedView(group, desiredTab);
                                if (null != newView)
                                {
                                    newView.IsSelected = true;
                                }
                            }
                        }
                        break;
                    }
            }
        }

        /// <summary>
        /// Obtain unpinned view at index. 
        /// </summary>
        public View GetUnpinnedView(ViewGroup viewGroup, int index)
        {
            return GetView(viewGroup, index, view => !view.IsPinned);
        }

        /// <summary>
        /// Obtain pinned view at index.
        /// </summary>
        public View GetPinnedView(ViewGroup viewGroup, int index)
        {
            return GetView(viewGroup, index, view => view.IsPinned);
        }

        /// <summary>
        /// Returns the <paramref name="index"/>'th visible view matching <paramref name="predicate"/>
        /// </summary>
        View GetView(ViewGroup viewGroup, int index, Func<View, bool> predicate)
        {
            return viewGroup.VisibleChildren.OfType<View>().Where(predicate).Skip(index).FirstOrDefault();
        }
    }

    static class PkgCmdIDList
    {
        public const int DocWellContextMenu = 0x100;
        public const int DocWellContextGroup = 0x101;
        public const int DocWellGeneralCommandsGroup = 0x102;

        public const int cmdidToolsOptions = 0x104;

        public const int cmdidGoToPinnedTab1 = 0x108;
        public const int cmdidGoToPinnedTab2 = 0x109;
        public const int cmdidGoToPinnedTab3 = 0x10a;
        public const int cmdidGoToPinnedTab4 = 0x10b;
        public const int cmdidGoToPinnedTab5 = 0x10c;
        public const int cmdidGoToPinnedTab6 = 0x10d;
        public const int cmdidGoToPinnedTab7 = 0x10e;
        public const int cmdidGoToPinnedTab8 = 0x10f;
        public const int cmdidGoToPinnedTab9 = 0x110;
        public const int cmdidGoToPinnedTab10 = 0x111;

        public const int cmdidGoToUnpinnedTab1 = 0x112;
        public const int cmdidGoToUnpinnedTab2 = 0x113;
        public const int cmdidGoToUnpinnedTab3 = 0x114;
        public const int cmdidGoToUnpinnedTab4 = 0x115;
        public const int cmdidGoToUnpinnedTab5 = 0x116;
        public const int cmdidGoToUnpinnedTab6 = 0x117;
        public const int cmdidGoToUnpinnedTab7 = 0x118;
        public const int cmdidGoToUnpinnedTab8 = 0x119;
        public const int cmdidGoToUnpinnedTab9 = 0x11a;
        public const int cmdidGoToUnpinnedTab10 = 0x11b;
    };
}
