using INFITF;
using KinTypeLib;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using Application = INFITF.Application;

namespace Automation-Functions_Methods
{
    public static class CreateJointsAndCommands
    {

        public static void CreateMechanismWithJoints(Application Catia, MechanismClass mechanismClass)
        {
            #region Required Instantiation

            CatiaJointMapping catiaJointMapping = new CatiaJointMapping();
            List<JointClass> joints = mechanismClass.Joints;
            ProductDocument productDocument = CatiaHandler.GetActiveProductDocument(Catia);
            Product mechanismProduct = CatiaHandler.GetProductFromProductDocument(productDocument);
            mechanismProduct.ApplyWorkMode(CatWorkModeType.DESIGN_MODE);  //set catia product to design mode
            string fixedPartWF = mechanismClass.FixedPartWF;
            string mechanismName = mechanismClass.Name;
            #endregion

            Product fixedPartProduct = CatiaAndReferenceHandling.GetProductByName(mechanismProduct, fixedPartWF);
            Mechanism mechanism = CatiaAndReferenceHandling.CreateCatiaMechansim(mechanismProduct, fixedPartProduct, mechanismName);

            #region "Creation of Joints & Commands"

            if (joints != null)
            {              
                foreach (JointClass joint in joints)
                {
                    CatiaAndReferenceHandling.FillProductInstancesByWireframeDefintion(joint, mechanismProduct); //filling the product instances again after deserializing since the reconstruction makes it null 
                    Object[] referenceObjectArray = CatiaAndReferenceHandling.CreateReferenceArrayFromJointClass(joint, mechanismProduct);
                    string internalName = catiaJointMapping.GetInternalJointNameByJointType(joint.Type);
                    Joint newJoint = null;
                    newJoint = mechanism.AddJoint(internalName, referenceObjectArray);
                    if (joint.Command != null)
                    {
                       CatiaAndReferenceHandling.CreateCommandFromJointClass(joint, newJoint, mechanism);
                    }
                }
            }
            else
            {
                MessageBox.Show("No joints have been created! Something's wrong with the Excel!");
            }
            #endregion
        }
    }

}