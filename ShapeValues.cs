using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class ShapeValues
    {
        //for point
        private double xValue ;
        private double yValue ;
        private double zVAlue ;

        private double xMeasured;
        private double yMeasured;
        private double zMeasured;

        //for line
        private double beginOffset;
        private double endOffset;
        private string lengthType;


        [XmlAttribute]
        public double XValue { get => xValue; set => xValue = value; }

        [XmlAttribute]
        public double YValue { get => yValue; set => yValue = value; }

        [XmlAttribute]
        public double ZValue { get => zVAlue; set => zVAlue = value; }

        [XmlAttribute]
        public double XMeasured { get => xMeasured; set => xMeasured = value; }

        [XmlAttribute]
        public double YMeasured { get => yMeasured; set => yMeasured = value; }

        [XmlAttribute]
        public double ZMeasured { get => zMeasured; set => zMeasured = value; }

        [XmlAttribute]
        public double BeginOffset { get => beginOffset; set => beginOffset = value; }

        [XmlAttribute]
        public double EndOffset { get => endOffset; set => endOffset = value; }

        [XmlAttribute]
        public string LengthType { 
            get => lengthType;
            set
            {
                lengthType = value;
                InternalLengthType = GetTheCatiaLengthTypeInInteger(value);
            }
        }
        [XmlIgnore]
        public int InternalLengthType { get; set; }

        private  int GetTheCatiaLengthTypeInInteger(string lengthtype)
        {
            int lengthInInteger = 0;
            switch (lengthtype)
            {
                case "Length":
                    lengthInInteger = 0;
                    break;
                case "Infinite":
                    lengthInInteger = 1;
                    break;
                case "InfiniteStartPoint":
                    lengthInInteger = 2;
                    break;
                case "InfiniteEndPoint":
                    lengthInInteger = 3;
                    break;
                default:
                    break;
            }

            return lengthInInteger;
        }

    }
}