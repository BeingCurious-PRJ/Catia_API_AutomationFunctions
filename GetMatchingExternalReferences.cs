using INFITF;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation-Functions_Methods
{
    class GetMatchingExternalReferences
    {
        #region Catia Initials
        public static Application Catia = CatiaHandler.GetCatiaObject();
        public static ProductDocument document = CatiaHandler.GetActiveProductDocument(Catia);
        public static Product product = CatiaHandler.GetProductFromProductDocument(document);
        #endregion

        public static void MatchExternalReferences(MainGeometryClass mainGeometryParts, MainGeometryClass mainGeometrySkeleton)
        {
            ExternalReferenceMatchClass parts = new ExternalReferenceMatchClass()
            {
                PartElements = new List<PartElement>()
            };
            foreach (GeoPartClass geoPartClass in mainGeometryParts.PartClass)
            {
                PartElement partElement = new PartElement
                {
                    Name = geoPartClass.Definition,
                    ExternalReference = new List<ExternalReferences>()
                };
                foreach (Geometry geometry in geoPartClass.Geometries)
                {
                    if (geometry.HybridBodyName == "External References")
                    {
                        ExternalReferences externalReferences = new ExternalReferences
                        {
                            Point = geometry.Name,
                            XCoord = geometry.ShapeSpecific[0].XMeasured,
                            YCoord = geometry.ShapeSpecific[0].YMeasured,
                            ZCoord = geometry.ShapeSpecific[0].ZMeasured,
                        };
                        foreach (GeoPartClass partClass in mainGeometrySkeleton.PartClass)
                        {
                            foreach (Geometry geometry1 in partClass.Geometries)
                            {
                                if (geometry1.Type == "HybridShapePointCoord" || geometry1.Type == "HybridShapePointTangent")
                                {
                                    if (geometry.ShapeSpecific[0].XMeasured == geometry1.ShapeSpecific[0].XMeasured &&
                                        geometry.ShapeSpecific[0].YMeasured == geometry1.ShapeSpecific[0].YMeasured &&
                                        geometry.ShapeSpecific[0].ZMeasured == geometry1.ShapeSpecific[0].ZMeasured &&
                                        geometry.Name == geometry1.Name)
                                    {
                                        externalReferences.MatchPoint = geometry1.Name;
                                    }

                                }
                            }
                        }
                        partElement.ExternalReference.Add(externalReferences);
                    }
                }
                parts.PartElements.Add(partElement);
            }
            string xmlPath = DialogBoxHandling.OpenDialogBoxAndSaveXMLFilePath();
            Serializing.SerilizeExternalRefMatch(parts, xmlPath);
        }
    }

}