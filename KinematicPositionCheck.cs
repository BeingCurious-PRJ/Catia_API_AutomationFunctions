using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class KinematicPositionCheck
    {
        private List<KinMovement> kinMovements;

        public List<KinMovement> KinMovements { get => kinMovements; set => kinMovements = value; }

    }
}