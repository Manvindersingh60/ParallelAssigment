using ExcelDna.Integration;

namespace ParallelAssigment
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Utility.Form = new ProcessingForm();
            Utility.Form.Show();
            Utility.Form.Hide();
            ExcelAsyncUtil.Initialize();
        }

        public void AutoClose()
        {
        }
    }
}
