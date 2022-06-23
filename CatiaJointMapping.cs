using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Automation-Functions_Methods
{
    public class CatiaJointMapping
    {
        public CatiaJointMapping()//constructor without parameter...hence okay for serializing
        {
            CatiaJointInformations = OpenAndReadExcel();
        }
        private List<CatiaJointInformation> catiaJointInformations;
        public List<CatiaJointInformation> CatiaJointInformations { get => catiaJointInformations; set => catiaJointInformations = value; }
      
        private List<CatiaJointInformation> OpenAndReadExcel()
        {
            #region Opening the specific Reference file

            string excelFileName = "JOINTS_References";
            string excelSheetName = "CATJoints";
            string excelFileFullPath = ExcelHandling.GetFullRessourcePath(excelFileName);
            Excel.Worksheet worksheet = ExcelHandling.GetWorksheetFromWorkbook(excelSheetName, excelFileFullPath);
            int rows = ExcelHandling.GetLastRowFromWorksheet(worksheet); //total number of used rows
            int columns = ExcelHandling.GetLastColumFromWorksheet(worksheet); //total number of used columns
            #endregion

            int r;
            List<CatiaJointInformation> detailsOfJointsLists = new List<CatiaJointInformation>();          
            for (r = 2; r <= rows; ++r)
            {              
                List<string> referenceGeometryTypes = new List<string>();
                for (int c = 5; c <= columns; ++c)
                {
                    if ((string)(worksheet.Cells[r, c] as Excel.Range).Value != "-")
                        referenceGeometryTypes.Add((string)(worksheet.Cells[r, c] as Excel.Range).Value);
                }                               
                detailsOfJointsLists.Add(new CatiaJointInformation()
                {
                    JointTypeName = (string)(worksheet.Cells[r, 1] as Excel.Range).Value,
                    CatiaInternalName = (string)(worksheet.Cells[r, 3] as Excel.Range).Value,
                    CommandName = (string)(worksheet.Cells[r, 4] as Excel.Range).Value,
                    ReferenceGeometryType = referenceGeometryTypes
                }) ;
            }
            return detailsOfJointsLists;
        }
        public string GetInternalJointNameByJointType(string jointTypeName)
        {
            string result = "";
            foreach (CatiaJointInformation item in catiaJointInformations)
            {
                if (item.JointTypeName == jointTypeName)
                {
                    result = item.CatiaInternalName;
                }
            }
            return result;
        }
        public int GetNumberOfInputsbyJointType(string jointTypeName)
        {
            int result = -1;
            foreach (CatiaJointInformation item in catiaJointInformations)
            {
                if (item.JointTypeName == jointTypeName)
                {
                    result = item.ReferenceGeometryType.Count;
                }
            }
            return result;
        }         
    }

}