using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Automation-Functions_Methods
{
    public class CatiaJointInformation
    {
        private string jointTypeName;
        private string catiaInternalName;
        private string commandName;
        private List<string> referenceGeometryType;

        public string JointTypeName { get => jointTypeName; set => jointTypeName = value; }
        public string CatiaInternalName { get => catiaInternalName; set => catiaInternalName = value; }
        public string CommandName { get => commandName; set => commandName = value; }
        public List<string> ReferenceGeometryType { get => referenceGeometryType; set => referenceGeometryType = value; }
    }

}