using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class JointClass
    {
         #region Fields
        private string name;
        private string type;
        private string internalname;
        private CatiaCommand command;
        private List<JointPart> jointParts;
        private int numberofInputs;
        #endregion

        #region Properties
        [XmlAttribute("Name")]
        public string Name { get => name; set => name = value; }

        [XmlAttribute("Type")]
        public string Type
        {
            get => type;
            set
            {
                type = value;
                var getNoOfInputsAndCatiaInternalName = GetNoOfInputsAndCatiaInternalName(type);
                NumberofInputs = getNoOfInputsAndCatiaInternalName.NumberOfInputs;
                Internalname = getNoOfInputsAndCatiaInternalName.InternalCatiaJointName;
            }
        }

        [XmlAttribute("NumberOfInputs")]
        public int NumberofInputs { get => numberofInputs; set => numberofInputs = value; }

        [XmlAttribute("InternalName")]
        public string Internalname { get => internalname; set => internalname = value; }

        public List<JointPart> JointParts { get => jointParts; set => jointParts = value; }
        public CatiaCommand Command { get => command; set => command = value; }
        #endregion

        private (int NumberOfInputs, string InternalCatiaJointName) GetNoOfInputsAndCatiaInternalName(string type)
        {
            numberofInputs = 0;
            internalname = "";
            if (Type != null)
            {
                switch (Type.ToLower())
                {
                    case "revolute":
                        numberofInputs = 4;
                        internalname = "CATKinRevoluteJoint";
                        break;

                    case "prismatic":
                        numberofInputs = 4;
                        internalname = "CATKinPrismaticJoint";
                        break;

                    case "spherical":
                        numberofInputs = 2;
                        internalname = "CATKinSphericalJoint";
                        break;

                    case "u joint":
                        numberofInputs = 2;
                        internalname = "CATKinUJoint";
                        break;

                    case "cylindrical":
                        numberofInputs = 2;
                        internalname = "CATKinCylindricalJoint";
                        break;

                    case "point surface":
                        numberofInputs = 2;
                        internalname = "CATKinPointSurfaceJoint";
                        break;

                    case "point curve":
                        numberofInputs = 2;
                        internalname = "CATKinPointCurveJoint";
                        break;

                    case "rigid":
                        numberofInputs = 2;
                        internalname = "CATKinRigidJoint";
                        break;

                    default:
                        break;
                }
            }
            return (numberofInputs, internalname);
        }
        public JointClass()
        {
            this.Command = null;
        }

    }
}