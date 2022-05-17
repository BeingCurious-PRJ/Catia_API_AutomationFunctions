using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automation_Functions_Methods
{
    public static class WindowsDialogBox_Handling
    {
        static string xmlPath;
        public static string OpenDialogBoxAndGetExcelFilePath()
        {
            string excelPath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //Multiselect = true,
                //openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                Filter = "Excel Files|*.xls; *.xlsx"
            };// need "Microsoft.Win32" namespace to open the dialog box and select a file like in normal windows application 
            if(openFileDialog.ShowDialog() == true)
            {
                //foreach (string filename in openFileDialog.FileNames)
                //{
                //    mainWindow.Files.Items.Add(System.IO.Path.GetFileName(filename));
                //}
                excelPath = openFileDialog.FileName;
            }
            //string excelFileRequired = openFileDialog.SafeFileName;
            return excelPath;
        }
        public static Excel.Worksheet DropDownListSelectionOfWorksheetFromExcel(string worksheetName, string excelWorkBookPath)
        {
            
        }
        public static string OpenDialogBoxAndSaveXMLFilePath()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "XML Files(*.xml)|*.xml",
                FilterIndex = 1
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                xmlPath = saveFileDialog.FileName;
            }
            return xmlPath;
        }

        public static string OpenExisitingXMLFile()
        {
            string XMLPath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "XML Files(*.xml)|*.xml|All files (*.*)|*.*",
                FilterIndex = 1
            };
            if (openFileDialog.ShowDialog() == true)
            {
                XMLPath = openFileDialog.FileName;
            }
            return XMLPath;
        }
    }
}