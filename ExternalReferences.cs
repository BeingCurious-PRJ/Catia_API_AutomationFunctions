using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class ExternalReferences
    {
        private string point;
        private string matchPoint;
        private double xCoord;
        private double yCoord;
        private double zCoord;
        private bool exists;

        [XmlAttribute]
        public string Point { get => point; set => point = value; }
        [XmlAttribute]
        public string MatchPoint { get => matchPoint; set => matchPoint = value; }
        [XmlAttribute]
        public double XCoord { get => xCoord; set => xCoord = value; }
        [XmlAttribute]
        public double YCoord { get => yCoord; set => yCoord = value; }
        [XmlAttribute]
        public double ZCoord { get => zCoord; set => zCoord = value; }
        [XmlAttribute]
        public bool Exists { get => exists; set => exists = value; }

    }
}