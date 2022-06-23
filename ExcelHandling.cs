using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Automation_Functions_Methods
{
    public static class ExcelHandling
    {
        private static Excel.Application application;
        private static readonly string directoryForFile = Environment.CurrentDirectory;
        private static readonly string directoryFile = Directory.GetParent(directoryForFile).Parent.FullName;

        public static Excel.Application Application 
        { 
            get
            {
                if (application == null)
                {
                    application = new Excel.Application();
                }
                return application;
            }           
            set => application = value; 
        }
        public static string GetFullRessourcePath(string fileName)  //only used to import 'Joints References' excel file from the Ressource folder##used in CatiaJointMapping class 
        {
            return directoryFile + @"\Ressources\" + fileName;
        }

        #region Getting Workbook and the worksheet

        public static Workbook GetWorkbook(string filePath)
        {
            Workbook workbook = null;
            try
            {
                workbook = Application.Workbooks.Open(filePath);
            }
            catch (FileNotFoundException ex)
            {
                var message = MessageBox.Show($"Workbook '{filePath}' could not be opened!");
                throw ex;
            }
            return workbook;
        }
        public static Worksheet GetWorksheet(string worksheetName, Workbook workbook)
        {
            Worksheet worksheet = null;
            try
            {
                worksheet = (Worksheet)workbook.Worksheets[worksheetName];
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"Worksheet '{worksheetName}' in '{workbook.Name}'could not be found!");
                throw ex;
            }

            return worksheet;
        }
        public static Worksheet GetWorksheetFromWorkbook (string worksheetName, string workbookPath)
        {
            Workbook workbook = GetWorkbook(workbookPath);
            Worksheet worksheet = null;
            if (workbook != null)
            {
                worksheet = GetWorksheet(worksheetName, workbook);
            }           
            return worksheet;
        }

        public static List<string> GetListOfWorksheetnamesFromWorkboot(Workbook workbook)
        {
            List<string> result = new List<string>();

            foreach (Worksheet item in workbook.Worksheets)
            {
                result.Add(item.Name);
            }


            return result;
        }
        #endregion

        #region Getting rows and columns

        private static Range GetUsedRangeFromWorksheet(Worksheet worksheet)
        { return worksheet.UsedRange; }

        private static int GetUsedRangeRow(Range usedRange)
        { return usedRange.Rows.Count; }

        private static int GetUsedRangeColumn(Range usedRange)
        { return usedRange.Columns.Count; }

        public static int GetLastRowFromWorksheet(Worksheet worksheet)
        {
            Range usedrange = GetUsedRangeFromWorksheet(worksheet);
            return GetUsedRangeRow(usedrange);
        }
        public static int GetLastColumFromWorksheet(Worksheet worksheet)
        {
            Range usedrange = GetUsedRangeFromWorksheet(worksheet);
            return GetUsedRangeColumn(usedrange);
        }
        #endregion

        #region Closing Excel Instances
        public static void CloseExcelInstance(Worksheet worksheet)
        {
            Excel.Application app = worksheet.Parent.Parent;
            Workbook wb = worksheet.Parent;
            wb.Close();
            Marshal.ReleaseComObject(wb);
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
        #endregion

    }


}