using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MECMOD;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class HybridBodyClass
    {
        private string name = "";
        private string parent = "";
        private string geoSetID = "";
        private string parentGeoSetID = "";
        private HybridBody hybridbodyObject = null;


        [XmlAttribute]
        public string Name { get => name; set => name = value; }

        [XmlAttribute]
        public string Parent { get => parent; set => parent = value; }

        [XmlAttribute]
        public string ParentGeoSetID { get => parentGeoSetID; set => parentGeoSetID = value; }

        [XmlAttribute]
        public string GeoSetID { get => geoSetID; set => geoSetID = value; }

        [XmlIgnore]
        public HybridBody HybridbodyObject { get => hybridbodyObject; set => hybridbodyObject = value; }

    }
}