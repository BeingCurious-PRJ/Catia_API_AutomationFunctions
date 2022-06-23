using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class JointPart
    {
        [XmlIgnore] //ignores the property during deserializing too and gives null after reconstructing--we need ProductInstanceName after deserializing to correct it for creating references and hence joints :|
        public string ProductInstanceName;

        [XmlAttribute("WireframeDefinition")]
        public string WireframeDefinition;

        [XmlAttribute("InstanceNumber")]
        public string PartInstanceNumber;

        public List<InputGeometry> InputGeometries;

        public JointPart()
        {
            InputGeometries = new List<InputGeometry>();
        }
        //[System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        //public bool ShouldSerializeProductInstanceName() { return false; } //This will cause the ProductInstanceName to not be serialized, but still allow it to be deserialized####Not working.
        

    }
}