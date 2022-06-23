using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class PartElement
    {
        private string name;
        private List<ExternalReferences> externalReference;

        [XmlAttribute]
        public string Name { get => name; set => name = value; }

        public List<ExternalReferences> ExternalReference { get => externalReference; set => externalReference = value; }

    }
}