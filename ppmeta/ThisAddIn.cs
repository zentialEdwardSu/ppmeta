using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ppmeta
{
    internal static class ShareState
    {
        public static Config Config { get; set; }
        public static string Code { get; set; }
        public static List<int> TrackedSlideIds { get; set; } = new List<int>();
        public static string TrackedSourceCode { get; set; } = string.Empty;
        public static bool IsTracking { get; set; } = false;
        
        public static void ClearTracking()
        {
            TrackedSlideIds.Clear();
            TrackedSourceCode = string.Empty;
            IsTracking = false;
        }
    }
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ShareState.Config = Config.Load();
            ShareState.Code = string.Empty;
            //this.Application.PresentationNewSlide += new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        /// <summary>
        ///  add a text box to the new slide when a new slide is created
        /// </summary>
        /// <param name="Sld"></param>
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }


        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
