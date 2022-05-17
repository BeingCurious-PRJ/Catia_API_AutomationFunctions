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
    }
    public static string GetFullResourcePath(string fileName)  //only used to import 'specific' excel file from "a folder -FolderName within the repo"
    {
        return directoryFile + @"\FolderName\" + fileName;
    }

    #region  "Getting Workbook and Worksheets from ExcelFile"
    public static Workbook GetWorkbook(string filePath)
    {
        Workbook workbook = null;
        try
        {
            workbook = Application.Workbooks.Open(filePath);
        }
        catch(FileNotFoundException ex)
        {
            var message = MessageBox.Show($"Workbook '{filePath}' could not be opened!");
            throw ex;
        }
        return workbook;
    }
    public static Worksheet GetWorksheet(string worksheetName, Workbook workbook)
    {
        worksheetName worksheet = null;
        try
        {
            worksheet = (Worksheet)workbook.Worksheets[worksheetName];
        }
        catch(Exception ex)
        {
            _ = MessageBox.Show($"Worksheet '{worksheetName}' in '{workbook.Name}'could not be found!");
            throw ex;  
        }
        return worksheet;
    }

    #endregion



}