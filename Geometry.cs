using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MECMOD;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class Geometry
    {
        private string name = "";
        private string type = "";
        private string hybridBodyName = "";
        private string geoSetID = "";
        private string elementID = "";
        private HybridShape hybridShapeObject;

        private List<ShapeValues> shapeSpecific;
        private List<ParentGeometry> parentGeometries;


        [XmlAttribute]
        public string Name { get => name; set => name = value; }

        [XmlAttribute]
        public string Type { get => type; set => type = value; }

        [XmlAttribute]
        public string HybridBodyName { get => hybridBodyName; set => hybridBodyName = value; }

        [XmlAttribute]
        public string GeoSetID { get => geoSetID; set => geoSetID = value; }

        [XmlAttribute]
        public string ElementID { get => elementID; set => elementID = value; }

        public List<ShapeValues> ShapeSpecific { get => shapeSpecific; set => shapeSpecific = value; }
        public List<ParentGeometry> ParentGeometries { get => parentGeometries; set => parentGeometries = value; }
        
        [XmlIgnore]
        public HybridShape HybridShapeObject { get => hybridShapeObject; set => hybridShapeObject = value; }

    }
}