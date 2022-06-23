using ACEFramework;
using INFITF;
using KinTypeLib;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
//using System.Linq;
//using ACEBaseLib;

namespace Automation-Functions_Methods
{
public static class CatiaAndReferenceHandling
    {        
        #region "Reference Handling"
        public static Object[] CreateReferenceArrayFromJointClass(JointClass joint, Product mechansimProduct)
        {
            List<string> stringList = GetGeometryPathList(joint);
            List<object> objectList = GetReferenceObjectListFromPathStringList(stringList, mechansimProduct);
            return (Object[])objectList.ToArray();
        }
        public static List<string> GetGeometryPathList(JointClass joint)
        {
            List<string> list = new List<string>();
            JointPart jointPart1 = joint.JointParts[0];
            JointPart jointPart2 = joint.JointParts[1];
            if (joint.Internalname == "CATKinRigidJoint")
            {
                jointPart1.InputGeometries[0].GeometryPath = $"{jointPart1.ProductInstanceName}/!";
                jointPart2.InputGeometries[0].GeometryPath = $"{jointPart2.ProductInstanceName}/!";
                list.Add(jointPart1.InputGeometries[0].GeometryPath);
                list.Add(jointPart2.InputGeometries[0].GeometryPath);
            }
            else
            {
                for (int i = 0; i < jointPart1.InputGeometries.Count; i++)
                {
                    jointPart1.InputGeometries[i].GeometryPath = $"{jointPart1.ProductInstanceName}/!{jointPart1.InputGeometries[i].GeometryPath}";
                    jointPart2.InputGeometries[i].GeometryPath = $"{jointPart2.ProductInstanceName}/!{jointPart2.InputGeometries[i].GeometryPath}";
                    list.Add(jointPart1.InputGeometries[i].GeometryPath);
                    list.Add(jointPart2.InputGeometries[i].GeometryPath);
                }
            }
            return list;
        }
        public static List<object> GetReferenceObjectListFromPathStringList(List<string> stringList, Product mechanismProduct)
        {
            List<object> resultList = new List<object>();
            List<Reference> refList = new List<Reference>();
            foreach (string path in stringList)
            {
                Reference reference = mechanismProduct.CreateReferenceFromName(path);
                refList.Add(reference);
            }
            foreach (var item in refList)
            {
                resultList.Add((object)item);
            }
            return resultList;
        }
        public static string CreateReferenceStringForXML(string referenceString)
        {
            string endString;
            string referencePath;

            if (referenceString == null) // then it's probably a rigid
            {
                referencePath = "";
            }
            else
            {
                string[] splitTheString = referenceString.Split('/');
                if (Equals(splitTheString[splitTheString.Length - 1], "Edge"))
                {
                    endString = String.Join("/", splitTheString.Take(splitTheString.Length - 1));
                    var endpath = endString.Split('/');
                    string pathToTake = String.Join("/", endpath.Skip(1));
                    referencePath = pathToTake;
                }
                else
                {                   
                    string pathToTake = String.Join("/", splitTheString.Skip(1));
                    referencePath = pathToTake;
                }
            }
            return referencePath;
        }
        internal static void FillProductInstancesByWireframeDefintion(JointClass joint, Product product)
        {
            List<string> partInstanceNames = new List<string>();
            int numOfParts = product.Products.Count;
            JointPart jointPart1 = joint.JointParts[0];
            JointPart jointPart2 = joint.JointParts[1];
            for (int i = 1; i <= numOfParts; ++i)
            {
                partInstanceNames.Add(product.Products.Item(i).get_Name().ToString());
            }
            CheckingForCorrectInstanceName(product, partInstanceNames, jointPart1, jointPart2);
        }
        private static void CheckingForCorrectInstanceName(Product product, List<string> partInstanceNames, JointPart jointPart1, JointPart jointPart2)
        {
            for (int j = 0; j < partInstanceNames.Count; ++j)
            {
                string instantNumber = partInstanceNames[j].Split('.').Last();
                Product foundInstance = product.Products.Item(partInstanceNames[j]);
                if (Equals(foundInstance.get_Definition(), jointPart1.WireframeDefinition) && Equals(jointPart1.PartInstanceNumber, instantNumber))
                {
                    jointPart1.ProductInstanceName = $"{product.get_Name()}/{partInstanceNames[j]}";
                }
                if (Equals(foundInstance.get_Definition(), jointPart2.WireframeDefinition) && Equals(jointPart2.PartInstanceNumber, instantNumber))
                {
                    jointPart2.ProductInstanceName = $"{product.get_Name()}/{partInstanceNames[j]}";
                }
            }
        }

        #endregion

        #region "Catia Handling"
        public static Product GetProductByName(Product product, string productWF)
        {
            Product result = null;
            //result = product.Products.Item(productName);
            foreach(Product productfixed in product.Products)
            {
                if (productfixed.get_Definition() == productWF)
                {
                    result = productfixed;
                }
            }
            return result;
        }
        public static Mechanism CreateCatiaMechansim(Product product, Product fixedPartProduct, string mechansimName)
        {
            Mechanism resultMechanism = null;
            Mechanisms mechanisms = (Mechanisms)product.GetTechnologicalObject("Mechanisms");
            resultMechanism = mechanisms.Add();
            resultMechanism.set_Name(mechansimName);
            resultMechanism.FixedPart = fixedPartProduct;
            return resultMechanism;
        }
        public static void CreateCommandFromJointClass(JointClass jointInformation, Joint catiaJoint, Mechanism mechanism)
        {
            if (Equals(jointInformation.Type, "cylindrical")) //check if type is cylindrical
            {
                MechanismCommand newCommand1 = mechanism.AddCommand("CATKinLengthCmd", catiaJoint);
                newCommand1.set_Name(jointInformation.Command.Cname);
                catiaJoint.LowerLimit1 = jointInformation.Command.Llimit1;
                catiaJoint.UpperLimit1 = jointInformation.Command.Ulimit1;

                MechanismCommand newCommand2 = mechanism.AddCommand("CATKinAngleCmd", catiaJoint);
                newCommand2.set_Name(jointInformation.Command.Cname);
                catiaJoint.LowerLimit2 = jointInformation.Command.Llimit2;
                catiaJoint.UpperLimit2 = jointInformation.Command.Ulimit2;
            }
            else
            {
                MechanismCommand newCommand = mechanism.AddCommand("CATKinLengthCmd", catiaJoint);
                newCommand.set_Name(jointInformation.Command.Cname);
                catiaJoint.LowerLimit1 = jointInformation.Command.Llimit1;
                catiaJoint.UpperLimit1 = jointInformation.Command.Ulimit1;
            }
        }
        public static void GetCommandLimitsFromCatiaJoint(JointClass jointClass, CatiaCommand catiaCommand, Product product)
        {
            Application oCatia = (Application)Marshal.GetActiveObject("CATIA.Application");
            KinematicTools kinematicTools = (KinematicTools)oCatia.GetItem("ACEKinematicTools");
            Joint catiaJoint1 = (Joint)kinematicTools.getKinJointByCmdName(product, jointClass.Command.Cname);
            if (Equals(jointClass.Type, "cylindrical")) //check if type is cylindrical
            {
                catiaCommand.Llimit1 = catiaJoint1.LowerLimit1;
                catiaCommand.Ulimit1 = catiaJoint1.UpperLimit1;
                catiaCommand.Llimit2 = catiaJoint1.LowerLimit2;
                catiaCommand.Ulimit2 = catiaJoint1.UpperLimit2;
            }
            else
            {
                catiaCommand.Llimit1 = catiaJoint1.LowerLimit1;
                catiaCommand.Ulimit1 = catiaJoint1.UpperLimit1;
            }
        }
        #endregion

        #region "Wireframe Definition"... Handling U Joint

        public static void BringUjointPartsInCorrectOrder(JointClass joint)
        {
            string WF1 = joint.JointParts[0].WireframeDefinition;
            string WF2 = joint.JointParts[1].WireframeDefinition;
            int iwf1 = GetIndexFromDefinitionList(WF1, DefinitionList);
            int iwf2 = GetIndexFromDefinitionList(WF2, DefinitionList);
            if (iwf1 > iwf2 || (iwf1 == -1 && iwf2 != -1))
            {
                _ = ReverseJointPartOrder(joint);
            }
        }
        public static string GetWireframeDefinitionAndInstanceNumberByInstanceName(string productInstanceName, Product mechanismProduct, JointPart jointPart)
        {
            string exactInstanceName = productInstanceName.Split('/').Last();
            Product foundInstance = mechanismProduct.Products.Item(exactInstanceName);
            var instanceNumber = exactInstanceName.Split('.').Last();
            jointPart.PartInstanceNumber = instanceNumber;
            string result = foundInstance.get_Definition();
            return result;//return the WF definition
        }
        private static List<string> DefinitionList = new List<string>    //for pattern matching with *, C# provides a class termed as Regex which can be found in (System.Text.RegularExpression)
        {
            "*WF KAROSSERIE*",
            "*ACHSTRAEGER*",
            "*STUETZLAGER*",
            "*ZAHNSTANGE*",
            "*WF ZS LAGER*",
            "*WF UNIVERSAL ELEMENT*", // not sure about that
            "*RADTRAEGER*",
            "*SCHWENKLAGER*",
            "*STURZ EINSTELLUNG*",
            "*SPUR EINSTELLUNG*",
            "*LENKER*",
            "*ZS DS*",
            "*WF QL*",
            "*LAENGE*",
            "*AUSGLEICH LAENGE*",
            "*AUSGLEICH ABTRIEBSWELLEN*",
            "*ABTRIEBSWELLE*",
            "*DAEMPFER OBEN*",
            "*DAEMPFER UNTEN*",
            "*FEDER OBEN",
            "*FEDER UNTEN",
            "*FEDERBEIN UNTEN*",
            "*FEDERBEIN OBEN*",
            "*STABI*",
            "*PENDELSTUETZE OBEN*",
            "*PENDELSTUETZE UNTEN*",
            "*PENDELSTUEZE UNTEN*", // I kid u f**king not >_<
            "*PENDELSTUETZE*",
            "*WF HOEHENSTANDSSENSOR*",
            "*HOEHENSTANDSENSOR HEBEL*",
            "*HSS HEBEL*",
            "*HEBELARM HSS*",
            "*ANLENKSTANGE HSS*",
            "*HOEHENSTANDSENSOR PENDEL*",
            "*WF SPREIZACHSE*"
        };
        private static int GetIndexFromDefinitionList(string wireframeDefintion, List<string> definitionList)
        {
            int result = -1;
            for (int i = 0; i < definitionList.Count; i++)
            {
                if (Operators.LikeString(wireframeDefintion, definitionList[i], CompareMethod.Text))
                {
                    result = i;
                    break;
                }
            }
            return result;
        }
        private static JointClass ReverseJointPartOrder(JointClass jointClass)
        {
            JointPart swap = jointClass.JointParts[1];
            jointClass.JointParts[1] = jointClass.JointParts[0];
            jointClass.JointParts[0] = swap;
            return jointClass;
        }
        #endregion

    }

}