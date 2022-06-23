using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class GeoPartClass
    {
        private string partName = "";
        private string definition = "";
        private List<HybridBodyClass> hybridbodies;
        private List<Geometry> geometries;


        [XmlAttribute]
        public string PartName { get => partName; set => partName = value; }

        [XmlAttribute]
        public string Definition { get => definition; set => definition = value; }

        public List<HybridBodyClass> Hybridbodies{ get => hybridbodies; set => hybridbodies = value; }
        public List<Geometry> Geometries { get => geometries; set => geometries = value; }

    }
}