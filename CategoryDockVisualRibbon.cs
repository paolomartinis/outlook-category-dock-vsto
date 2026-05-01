using System;
using Microsoft.Office.Tools.Ribbon;

namespace CategoryDockVsto
{
    public partial class CategoryDockVisualRibbon : RibbonBase
    {
        private RibbonTab mailTab;
        private RibbonGroup categoryDockGroup;
        private RibbonButton openButton;

        public CategoryDockVisualRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            mailTab = Factory.CreateRibbonTab();
            categoryDockGroup = Factory.CreateRibbonGroup();
            openButton = Factory.CreateRibbonButton();

            mailTab.ControlId.ControlIdType = RibbonControlIdType.Office;
            mailTab.ControlId.OfficeId = "TabMail";
            mailTab.Label = "Home";

            categoryDockGroup.Label = "Category Dock";

            openButton.Label = "Category Dock";
            openButton.OfficeImageId = "CategorizeMenu";
            openButton.ShowImage = true;
            openButton.Click += OpenButton_Click;

            categoryDockGroup.Items.Add(openButton);
            mailTab.Groups.Add(categoryDockGroup);
            Tabs.Add(mailTab);
        }

        private void OpenButton_Click(object sender, RibbonControlEventArgs e)
        {
            Logger.Write("Visual ribbon button click.");
            Globals.ThisAddIn.ShowCategoryDock();
        }
    }
}
