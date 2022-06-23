using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class MechanismClass
    {
        #region Fields
        private string name;
        private string fixedPartWF;
        List<JointClass> joints;
        #endregion

        #region Properties

        [XmlAttribute("MechanismName")]
        public string Name { get => name; set => name = value; }

        [XmlAttribute("FixedPart")]
        public string FixedPartWF { get => fixedPartWF; set => fixedPartWF = value; }

        public  List<JointClass> Joints { get => joints; set => joints = value; }
        #endregion

        public (string name, string fixedPartName, List<JointClass> joints) GetMechanismClassInfo(string name, string fixedPartName, List<JointClass> joints)
        {
            Joints = joints ?? new List<JointClass>();
            Name = name;
            FixedPartWF = fixedPartName;

            return (Name, FixedPartWF, Joints);
        }

    }
}