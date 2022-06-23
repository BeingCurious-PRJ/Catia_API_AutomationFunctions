using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class ExternalReferenceMatchClass
    {
        private List<PartElement> partElements;

        public List<PartElement> PartElements { get => partElements; set => partElements = value; }

    }
}