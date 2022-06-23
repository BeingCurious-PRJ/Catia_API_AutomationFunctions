using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MECMOD;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class ParentGeometry
    {
        private  string name = "";
        private string elementID = "";
        private HybridShape hybridShapeObject;


        [XmlAttribute]
        public  string Name { get => name; set => name = value; }

        [XmlAttribute]
        public string ElementID { get => elementID; set => elementID = value; }
        [XmlIgnore]
        public HybridShape HybridShapeObject { get => hybridShapeObject; set => hybridShapeObject = value; }

    }
}