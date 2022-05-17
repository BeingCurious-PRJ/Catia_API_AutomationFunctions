using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Xml;  //contains classes to read and write xml documents
using System.Xml.Serialization;  //contains classes that are used to serialize objects into XML format documents or streams

namespace Automation_Functions_Methods
{
    public class XML_Serializing_Deserializing
    {
         #region Serializing
         public static string SerializeAnyClass( /*Class classobject,*/ string xmlFilePath)
         {
             XmlSerializer xmlSerializer = new XmlSerializer(/*typeof(Class)*/);
             XmlTextWriter textWriter = new XmlTextWriter(xmlFilePath, null);
             using (textWriter)
             {
                 xmlSerializer.Serialize(textWriter /*, classobject*/);
             }
             return xmlFilePath;
         }
        #endregion

        #region Deserializing

        public static Object DeserializeAnyClass(string xmlFilePathSaved)
        {
            object obj = null;
            XmlSerializer xmlDeserializer = new XmlSerializer(/*typeof(Class)*/);
            TextReader textReader = new StreamReader(xmlFilePathSaved);
            using (textReader)
            {
                obj = xmlDeserializer.Deserialize(textReader);
            }
            return obj;
        }
        #endregion
    }
}