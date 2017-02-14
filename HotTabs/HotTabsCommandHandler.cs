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
    internal sealed class HotTabsCommandHandler
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
        public static HotTabsCommandHandler Instance
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
            Instance = new HotTabsCommandHandler(package);
        }
        
        /// <summary>
                 /// Initializes a new instance of the <see cref="HotTabsCommandHandler"/> class.
                 /// Adds our command handlers for menu (commands must exist in the command table file)
                 /// </summary>
                 /// <param name="package">Owner package, not null.</param>
        private HotTabsCommandHandler(Package package)
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
}
