using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using INFITF;
using ProductStructureTypeLib;
using Application = INFITF.Application;

namespace Automation-Functions_Methods
{
    public static class ReadFromExcel
    {
        public static MechanismClass GetExcelContent(Application Catia, CatiaJointMapping catiaJointMapping, Excel.Worksheet ws, string xmlPath)
        {
            #region "Instantiation of required Variables/objects"

            ProductDocument productDocument = CatiaHandler.GetActiveProductDocument(Catia);
            Product mechanismProduct = CatiaHandler.GetProductFromProductDocument(productDocument);
            int rows = ExcelHandling.GetLastRowFromWorksheet(ws); //total number of used rows
            string mechanism_name = (string)(ws.Cells[1, 2] as Excel.Range).Value;
            string fixedpart_name = (string)(ws.Cells[7, 2] as Excel.Range).Value;
            List<JointClass> joints_list = new List<JointClass>(); //creating a list 'joints_list' to add objects of class 'JointClass' 
            MechanismClass mechanism = new MechanismClass();
            mechanism.GetMechanismClassInfo(mechanism_name, fixedpart_name, joints_list);
            #endregion

            for (int i = 10; i < rows; i++) // iteration starts from 'Joint' datas
            {
                string teststring1 = (string)(ws.Cells[i, 1] as Excel.Range).Value;
                if (teststring1 != null)  //to skip empty row
                {
                    JointClass newJoint = new JointClass(); //creating an object 'newJoint' of class 'JointClass'                                     
                    newJoint.Name = (string)(ws.Cells[i, 1] as Excel.Range).Value;
                    newJoint.Type = (string)(ws.Cells[i, 3] as Excel.Range).Value;
                    string commandString;
                    if ((commandString = (string)(ws.Cells[i, 4] as Excel.Range).Value) != null)
                    {
                        CatiaCommand catiaCommand = new CatiaCommand
                        {
                            Getcommandstring = commandString
                        };
                        if (catiaCommand.Cname != null)
                        {
                            newJoint.Command = catiaCommand;
                        }
                        else
                        {
                            MessageBox.Show($"Something is wrong in Line {i}", "Oopsies");
                            return null;
                        }
                    }
                    joints_list.Add(newJoint);
                    List<JointPart> jointParts = new List<JointPart>();
                    for (int h = 0; h < 2; ++h) //for two parts associated with each joint
                    {
                        JointPart newpart = new JointPart(); //creating an object 'newpart' of class 'JointPart'
                        newpart.ProductInstanceName = (string)(ws.Cells[i, 5] as Excel.Range).Value;
                        int numberOfInputs = catiaJointMapping.GetNumberOfInputsbyJointType(newJoint.Type);
                        int end = numberOfInputs / 2;
                        for (int j = 0; j < end; ++j)
                        {
                            InputGeometry inputGeometry = new InputGeometry();
                            inputGeometry.GeometryPath = (string)(ws.Cells[i, 6] as Excel.Range).Value;
                            if((string)(ws.Cells[i, 7] as Excel.Range).Value != null)
                            {
                                inputGeometry.GeometryType = ((string)(ws.Cells[i, 7] as Excel.Range).Value).Trim();
                            }
                            else
                            {
                                inputGeometry.GeometryType = (string)(ws.Cells[i, 7] as Excel.Range).Value;
                            }
                            newpart.InputGeometries.Add(inputGeometry);
                            i++;
                        }
                        jointParts.Add(newpart);
                    }
                    newJoint.JointParts = jointParts;
                }
            }
            PrepareTheExcelForSerializing(joints_list, mechanismProduct);
            Serializing.SerializeMechanism(mechanism, xmlPath);
            ExcelHandling.CloseExcelInstance(ws);
            return mechanism;
        }

        #region "Methods for Excel re-ordering(for serialization)"

        private static void PrepareTheExcelForSerializing(List<JointClass> joints_list, Product mechanismProduct)
        {
            foreach (JointClass newJoint in joints_list)
            {
                CatiaCommand catiaCommand1 = newJoint.Command;
                JointPart jointPart1 = newJoint.JointParts[0];
                JointPart jointPart2 = newJoint.JointParts[1];
                for (int i = 0; i < jointPart1.InputGeometries.Count; i++)
                {
                    jointPart1.WireframeDefinition = CatiaAndReferenceHandling.GetWireframeDefinitionAndInstanceNumberByInstanceName(jointPart1.ProductInstanceName, mechanismProduct,jointPart1);
                    jointPart1.InputGeometries[i].GeometryPath = CatiaAndReferenceHandling.CreateReferenceStringForXML(jointPart1.InputGeometries[i].GeometryPath);
                    
                    jointPart2.WireframeDefinition = CatiaAndReferenceHandling.GetWireframeDefinitionAndInstanceNumberByInstanceName(jointPart2.ProductInstanceName, mechanismProduct,jointPart2);
                    jointPart2.InputGeometries[i].GeometryPath = CatiaAndReferenceHandling.CreateReferenceStringForXML(jointPart2.InputGeometries[i].GeometryPath);
                }
                if (Equals(newJoint.Internalname, "CATKinUJoint"))
                {
                   CatiaAndReferenceHandling.BringUjointPartsInCorrectOrder(newJoint); // change the order from wireframe definition
                }
                if (newJoint.Command != null && newJoint.Command.Cmd == "" )
                {
                   CatiaAndReferenceHandling.GetCommandLimitsFromCatiaJoint(newJoint, catiaCommand1, mechanismProduct);
                }
                else if(newJoint.Command != null)
                {
                    if (Equals(newJoint.Type, "cylindrical")) //check if type is cylindrical
                    {
                        catiaCommand1.Llimit1 = newJoint.Command.Llimit1;
                        catiaCommand1.Ulimit1 = newJoint.Command.Ulimit1;

                        catiaCommand1.Llimit2 = newJoint.Command.Llimit2;
                        catiaCommand1.Ulimit2 = newJoint.Command.Ulimit2;
                    }
                    else
                    {
                        catiaCommand1.Llimit1 = newJoint.Command.Llimit1;
                        catiaCommand1.Ulimit1 = newJoint.Command.Ulimit1;
                    }
                }
            }
        }

        #endregion

    }

}