using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace ComInterop
{
    public class AddIn : IExcelAddIn
    {
        Application xlApp;
        public void AutoOpen()
        {
            xlApp = ExcelDnaUtil.Application as Application;
            Debug.Print(xlApp.Version);
            xlApp.WorkbookActivate += Wb => Debug.Print($"WorkbookActivate: {Wb.Name}");
        }

        public void AutoClose() { }
    }
}