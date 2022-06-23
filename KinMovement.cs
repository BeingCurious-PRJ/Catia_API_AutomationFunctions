using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class KinMovement
    {
        private string wF_Definition;
        private List<double[]> iAxisComponentArray;

        [XmlAttribute]
        public string WF_Definition { get => wF_Definition; set => wF_Definition = value; }
        [XmlArrayItem]
        public List<double[]> IAxisComponentArray { get => iAxisComponentArray; set => iAxisComponentArray = value; }

    }
}