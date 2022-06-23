using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Automation-Functions_Methods
{
    public class CatiaCommand
    {
        #region Fields
        private string cname;
        private string cmd;
        //private string direction;
        private double llimit1;
        private double ulimit1;
        private double llimit2;
        private double ulimit2;
        private string getcommandstring;
        #endregion

        #region Properties
        [XmlAttribute("Name")]
        public string Cname { get => cname; set => cname = value; }

        [XmlAttribute("Command")]
        public string Cmd { get => cmd; set => cmd = value; }

        //public string Direction { get => direction; set => direction = value; }

        [XmlAttribute("LowerLimit.1")]
        public double Llimit1 { get => llimit1; set => llimit1 = value; }

        [XmlAttribute("UpperLimit.1")]
        public double Ulimit1 { get => ulimit1; set => ulimit1 = value; }

        [XmlAttribute("LowerLimit.2")]
        public double Llimit2 { get => llimit2; set => llimit2 = value; }

        [XmlAttribute("UpperLimit.2")]
        public double Ulimit2 { get => ulimit2; set => ulimit2 = value; }

        [XmlIgnore]
        public string Getcommandstring
        { get => getcommandstring;
            set
            {
                getcommandstring = value;
                var getCommandString = GetCommandString(getcommandstring);
                Cname = getCommandString.Cname;
                Cmd = getCommandString.Cmd;
                Llimit1 = getCommandString.Llmit1;
                Ulimit1 = getCommandString.Ulimit1;
                Llimit2 = getCommandString.Llimit2;
                Ulimit2 = getCommandString.Ulimit2;
            }
        }
        #endregion

        private (string Cname, string Cmd, double Llmit1, double Ulimit1, double Llimit2, double Ulimit2) GetCommandString(string commandString)
        {
            string[] cinfo = commandString.Split(';');
            if (cinfo.Length ==1)
            {
                Cname = commandString;
                Cmd = "";
                Llimit1 = 0;
                Ulimit1 = 0;
                Llimit2 = 0;
                Ulimit2 = 0;
            }
            else if (cinfo.Length == 4)
            {
                Cname = cinfo[0];
                Cmd = cinfo[1];
                //Direction = cinfo[1];
                Llimit1 = Double.Parse(cinfo[2]);
                Ulimit1 = Double.Parse(cinfo[3]);
            }
            else if (cinfo.Length == 6)
            {
                Cname = cinfo[0];
                Cmd = cinfo[1];
                //Direction = cinfo[1];
                Llimit1 = Double.Parse(cinfo[2]);
                Ulimit1 = Double.Parse(cinfo[3]);
                Llimit2 = Double.Parse(cinfo[4]);
                Ulimit2 = Double.Parse(cinfo[5]);
            }

            return (Cname, Cmd, Llimit1, Ulimit1, Llimit2, Ulimit2);
        }

        //public CatiaCommand(string commandString)................//for serialization, we cannot have class constructor with parameters...so couldn't use you :(
        //{
        //    string[] cinfo = commandString.Split(';');
        //    if (cinfo.Length == 4)
        //    {
        //        Cname = cinfo[0];
        //        Cmd = cinfo[1];
        //        //Direction = cinfo[1];
        //        Llimit1 = Double.Parse(cinfo[2]);
        //        Ulimit1 = Double.Parse(cinfo[3]);
        //    }
        //    else if (cinfo.Length == 6)
        //    {
        //        Cname = cinfo[0];
        //        Cmd = cinfo[1];
        //        //Direction = cinfo[1];
        //        Llimit1 = Double.Parse(cinfo[2]);
        //        Ulimit1 = Double.Parse(cinfo[3]);
        //        Llimit2 = Double.Parse(cinfo[4]);
        //        Ulimit2 = Double.Parse(cinfo[5]);
        //    }

        //}

    }
}