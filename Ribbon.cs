using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace ComInterop
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
                  <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                  <ribbon>
                    <tabs>
                      <tab id='tab1' label='My Tab'>
                        <group id='group1' label='My Group'>
                          <button id='button1' label='My Button' onAction='OnButtonPressed'/>
                        </group >
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            var xlApp = ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application;
            var wb = xlApp.ActiveWorkbook;
            if (wb != null)
            {
                ShowMessage($"Active Workbook: {wb.Name}");
            }

            // ShowMessage("Hello from control " + control.Id);
        }

        void ShowMessage(string message)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                XlCall.Excel(XlCall.xlcAlert, message);
            });
        }
    }
}