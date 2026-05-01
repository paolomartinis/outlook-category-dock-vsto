using Microsoft.Office.Core;

namespace CategoryDockVsto
{
    public sealed class CategoryDockRibbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonId)
        {
            Logger.Write("GetCustomUI: " + ribbonId);
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""CategoryDockGroup"" label=""Category Dock"">
          <button id=""OpenCategoryDock""
                  label=""Category Dock""
                  size=""large""
                  imageMso=""CategorizeMenu""
                  onAction=""OpenCategoryDock""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(IRibbonUI ribbonUi)
        {
            Logger.Write("OnRibbonLoad.");
        }

        public void OpenCategoryDock(IRibbonControl control)
        {
            Logger.Write("OpenCategoryDock.");
            Globals.ThisAddIn.ShowCategoryDock();
        }
    }
}
