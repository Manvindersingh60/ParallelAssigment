using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ParallelAssigment
{
    /// <summary>
    /// Custom ribbon
    /// </summary>
    [ComVisible(true)]
    public class RequestCountRibbon : ExcelRibbon
    {
        /// <summary>
        /// Shows total number of api hits
        /// </summary>
        /// <param name="control1"></param>
        public void ShowRequestCount(IRibbonControl control1)
        {
            MessageBox.Show($"API hits: {Utility.SuccessfulHits}");
        }

        /// <summary>
        /// Generated ribbon UI xml
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='ParallelDots'>
            <group id='group1' label='Group1'>
              <button id='button1' imageMso='HappyFace' label='Show Count' onAction='ShowRequestCount'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }
    }
}
