using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class InputGeometry
    {
        private string geometryPath;
        private string geometryType;

        [XmlAttribute("GeometryPath")]
        public string GeometryPath { get => geometryPath; set => geometryPath = value; }

        [XmlAttribute("GeometryType")]
        public string GeometryType { get => geometryType; set => geometryType = value; }

        public bool ShouldSerializeGeometryPath()
        {
            return !String.IsNullOrEmpty(GeometryPath);
        }  //ignores the property "GeometryPath" whenever it is empty during serialization----z.B for Rigid Joint

    }
}