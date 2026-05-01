using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace CategoryDockVsto
{
    public partial class ThisAddIn
    {
        private CategoryDockForm paneControl;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private Timer startupTimer;
        private Office.CommandBarButton reopenButton;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Logger.Write("ThisAddIn_Startup.");
            AddReopenCommand();
            startupTimer = new Timer();
            startupTimer.Interval = 1200;
            startupTimer.Tick += StartupTimer_Tick;
            startupTimer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (reopenButton != null)
            {
                reopenButton.Click -= ReopenButton_Click;
                reopenButton = null;
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Logger.Write("CreateRibbonExtensibilityObject.");
            return new CategoryDockRibbon();
        }

        public void ShowCategoryDock()
        {
            Logger.Write("ShowCategoryDock.");
            if (taskPane == null)
            {
                paneControl = new CategoryDockForm(new CategoryService(Application));
                taskPane = CustomTaskPanes.Add(paneControl, "Category Dock");
                taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = 260;
            }

            taskPane.Visible = true;
        }

        private void StartupTimer_Tick(object sender, System.EventArgs e)
        {
            startupTimer.Stop();
            startupTimer.Dispose();
            startupTimer = null;
            ShowCategoryDock();
        }

        private void AddReopenCommand()
        {
            try
            {
                Office.CommandBars commandBars = Application.ActiveExplorer().CommandBars;
                Office.CommandBar standardBar = commandBars["Standard"];

                for (int i = standardBar.Controls.Count; i >= 1; i--)
                {
                    Office.CommandBarControl control = standardBar.Controls[i];
                    if (control.Tag == "CategoryDockVsto.Open")
                    {
                        control.Delete();
                    }
                }

                reopenButton = (Office.CommandBarButton)standardBar.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    System.Type.Missing,
                    System.Type.Missing,
                    System.Type.Missing,
                    true);
                reopenButton.Caption = "Category Dock";
                reopenButton.Tag = "CategoryDockVsto.Open";
                reopenButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                reopenButton.FaceId = 1295;
                reopenButton.Click += ReopenButton_Click;
                Logger.Write("Reopen command added.");
            }
            catch (System.Exception exception)
            {
                Logger.Write(exception);
            }
        }

        private void ReopenButton_Click(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            ShowCategoryDock();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
