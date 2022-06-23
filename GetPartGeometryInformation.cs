using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using System.Collections.Generic;
using HybridShapeTypeLib;


namespace Automation-Functions_Methods
{
   class GetPartGeometryInformation
    {
        #region Catia Initials
        public static Application Catia = CatiaHandler.GetCatiaObject();
        public static ProductDocument document = CatiaHandler.GetActiveProductDocument(Catia);
        public static Product product = CatiaHandler.GetProductFromProductDocument(document);
        #endregion
        public static int CatiaLengthType = 0;

        public static void GetPartDetailsAndSerialize(string partDefinition)
        {
            List<Part> partList = new List<Part>();
            MainGeometryClass mainGeometry = new MainGeometryClass
            {
                PartClass = new List<GeoPartClass>()
            };
            foreach (Product partitem in product.Products)
            {
                Part tempPart = (partitem.ReferenceProduct.Parent as PartDocument).Part;
                partList.Add(tempPart);
            }
            List<HybridBody> allHybridbodies = MethodsReadingGeometry.GetAllHybridbodiesFromParts(partList);

            foreach (Product part in product.Products)
            {
                GeoPartClass geoPartClass = new GeoPartClass();
                //geoPartClass = GetSkeletonInformation(allHybridbodies, part, geoPartClass);
                if (!string.IsNullOrEmpty(partDefinition))
                {
                    geoPartClass = GetSinglePartIformation(partDefinition, allHybridbodies, part, geoPartClass);
                }
                else
                {
                    geoPartClass = GetAllPartsInformation(allHybridbodies, part, geoPartClass);
                }
                mainGeometry.PartClass.Add(geoPartClass);
            }
            string geoxmlPath = DialogBoxHandling.OpenDialogBoxAndSaveXMLFilePath();
            Serializing.SerializeGeometry(mainGeometry, geoxmlPath);
        }

        private static GeoPartClass GetAllPartsInformation(List<HybridBody> allHybridbodies, Product part, GeoPartClass geoPartClass)
        {
            if (part.get_Definition() != "WF SKELETON" && part.get_Definition() != "" && part.get_Definition() != "SKELETON") //.............get all parts except skeleton
            {
                geoPartClass = DetailsForOtherParts(allHybridbodies, part);
            }

            return geoPartClass;
        }
        private static GeoPartClass GetSkeletonInformation(List<HybridBody> allHybridbodies, Product part, GeoPartClass geoPartClass)
        {
            if (part.get_Definition() == "WF SKELETON" || part.get_Definition() == "" || part.get_Definition() == "SKELETON") //.............for skeleton alone
            {
                geoPartClass = DetailsForSkeleton(allHybridbodies, part);
            }
            return geoPartClass;
        }
        private static GeoPartClass GetSinglePartIformation(string partDefinition, List<HybridBody> allHybridbodies, Product part, GeoPartClass geoPartClass)
        {
            if (part.get_Definition() == partDefinition && part.get_Definition() != "WF SKELETON") //............single part
            {
                geoPartClass = DetailsForOtherParts(allHybridbodies, part);
            }
            return geoPartClass;
        }

        private static GeoPartClass DetailsForSkeleton(List<HybridBody> allHybridbodies, Product part)
        {
            GeoPartClass partClass = new GeoPartClass();
            Part tempPart = (part.ReferenceProduct.Parent as PartDocument).Part;
            partClass.PartName = tempPart.get_Name();
            partClass.Hybridbodies = new List<HybridBodyClass>();
            partClass.Geometries = new List<Geometry>();
            foreach (HybridBody item in allHybridbodies)
            {
                //for fetching skeleton specific info...............................................................................................................
                if (item.get_Name() == "In and Out" || MethodsReadingGeometry.GetParentGeoSetFromChildGeoSets(item).get_Name() == "In and Out") //(item.get_Name() == "ALPHA MASTER Points" || MethodsReadingGeometry.GetParentGeoSetFromChildGeoSets(item).get_Name() == "ALPHA MASTER Points"
                        //|| item.get_Name() == "ALPHA MASTER Bauteile" || MethodsReadingGeometry.GetParentGeoSetFromChildGeoSets(item).get_Name() == "ALPHA MASTER Bauteile" || item.get_Name() == "Basics")
                {
                    Part parentPart = MethodsReadingGeometry.GetParentPartObjectFromAnyObject(item);
                    string parentName = parentPart.get_Name();
                    if (Equals(partClass.PartName, parentName))
                    {
                        HybridBodyClass hybridBodyClass = new HybridBodyClass
                        {
                            Name = item.get_Name(),
                            GeoSetID = item.GetHashCode().ToString()
                        };
                        if (item.Parent != parentPart)
                        {
                            HybridBody body = (HybridBody)item.Parent;
                            hybridBodyClass.Parent = body.get_Name();
                            foreach (HybridBodyClass hybridBody in partClass.Hybridbodies)
                            {
                                if (body.get_Name() == hybridBody.Name)
                                {
                                    hybridBodyClass.ParentGeoSetID = hybridBody.GeoSetID;
                                }
                            }
                        }
                        partClass.Hybridbodies.Add(hybridBodyClass);

                        List<HybridShape> hybridShape = new List<HybridShape>();
                        foreach (HybridShape shape in item.HybridShapes)
                        {
                            hybridShape.Add(shape);
                            Geometry geometry = new Geometry
                            {
                                HybridBodyName = item.get_Name(),
                                GeoSetID = hybridBodyClass.GeoSetID,
                                ElementID = shape.GetHashCode().ToString(),
                                Name = shape.get_Name()
                            };
                            List<AnyObject> inputList = MethodsReadingGeometry.GetInputParentsFromHybridshapes(shape, tempPart, geometry);
                            geometry.ParentGeometries = new List<ParentGeometry>();
                            if (inputList != null)
                            {
                                foreach (AnyObject hybrid in inputList)
                                {
                                    ParentGeometry parentGeometry = new ParentGeometry();
                                    if (hybrid != null)
                                    {
                                        parentGeometry.Name = hybrid.get_Name();
                                    }
                                    else
                                    {
                                        parentGeometry.Name = "";
                                    }
                                    geometry.ParentGeometries.Add(parentGeometry);
                                }
                            }
                            geometry.ShapeSpecific = new List<ShapeValues>();
                            if (shape is HybridShapePointCoord pointCoordinate)
                            {
                                ShapeValues shapeValues = new ShapeValues
                                {
                                    XValue = pointCoordinate.X.Value,
                                    YValue = pointCoordinate.Y.Value,
                                    ZValue = pointCoordinate.Z.Value
                                };
                                (shapeValues.XMeasured, shapeValues.YMeasured, shapeValues.ZMeasured) = MethodsReadingGeometry.MeasuringPointCoordinate(tempPart, shape, Catia);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            else if (shape is HybridShapeLinePtPt linePtPt)
                            {
                                ShapeValues shapeValues = new ShapeValues
                                {
                                    BeginOffset = linePtPt.BeginOffset.Value,
                                    EndOffset = linePtPt.EndOffset.Value
                                };
                                CatiaLengthType = linePtPt.GetLengthType();
                                shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            else if (shape is HybridShapeLinePtDir linePtDir)
                            {
                                ShapeValues shapeValues = new ShapeValues
                                {
                                    BeginOffset = linePtDir.BeginOffset.Value,
                                    EndOffset = linePtDir.EndOffset.Value
                                };
                                CatiaLengthType = linePtDir.GetLengthType();
                                shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            else if (shape is HybridShapeLineAngle lineAngle)
                            {
                                ShapeValues shapeValues = new ShapeValues
                                {
                                    BeginOffset = lineAngle.BeginOffset.Value,
                                    EndOffset = lineAngle.EndOffset.Value
                                };
                                CatiaLengthType = lineAngle.GetLengthType();
                                shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            else if (shape is HybridShapeLineNormal lineNormal)
                            {
                                ShapeValues shapeValues = new ShapeValues
                                {
                                    BeginOffset = lineNormal.BeginOffset.Value,
                                    EndOffset = lineNormal.EndOffset.Value
                                };
                                CatiaLengthType = lineNormal.GetLengthType();
                                shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            else if (geometry.Type == "") // WF Karosserie has external references without any specified type (but is a point)
                            {
                                geometry.Type = "HybridShapePointTangent";
                                ShapeValues shapeValues = new ShapeValues();
                                (shapeValues.XMeasured, shapeValues.YMeasured, shapeValues.ZMeasured) = MethodsReadingGeometry.MeasuringPointCoordinate(tempPart, shape, Catia);
                                geometry.ShapeSpecific.Add(shapeValues);
                            }
                            partClass.Geometries.Add(geometry);
                        }
                        MethodsReadingGeometry.AssignParentElementID(partClass);
                    }
                }
            }
            return partClass;
        }
        private static GeoPartClass DetailsForOtherParts(List<HybridBody> allHybridbodies, Product part)
        {
            Part tempPart = (part.ReferenceProduct.Parent as PartDocument).Part;
            GeoPartClass partClass = new GeoPartClass
            {
                PartName = tempPart.get_Name(),
                Definition = part.get_Definition(),
                Hybridbodies = new List<HybridBodyClass>(),
                Geometries = new List<Geometry>()
            };
            foreach (HybridBody item in allHybridbodies)
            {
                Part parentPart = MethodsReadingGeometry.GetParentPartObjectFromAnyObject(item);
                string parentName = parentPart.get_Name();
                if (Equals(partClass.PartName, parentName))
                {
                    HybridBodyClass hybridBodyClass = new HybridBodyClass
                    {
                        Name = item.get_Name(),
                        GeoSetID = item.GetHashCode().ToString()
                    };
                    if (item.Parent != parentPart)
                    {
                        HybridBody body = (HybridBody)item.Parent;
                        hybridBodyClass.Parent = body.get_Name();
                        foreach (HybridBodyClass hybridBody in partClass.Hybridbodies)
                        {
                            if (body.get_Name() == hybridBody.Name)
                            {
                                hybridBodyClass.ParentGeoSetID = hybridBody.GeoSetID;
                            }
                        }
                    }
                    partClass.Hybridbodies.Add(hybridBodyClass);

                    List<HybridShape> hybridShape = new List<HybridShape>();
                    foreach (HybridShape shape in item.HybridShapes)
                    {
                        hybridShape.Add(shape);
                        Geometry geometry = new Geometry
                        {
                            HybridBodyName = item.get_Name(),
                            GeoSetID = hybridBodyClass.GeoSetID,
                            ElementID = shape.GetHashCode().ToString(),
                            Name = shape.get_Name()
                        };
                        List<AnyObject> inputList = MethodsReadingGeometry.GetInputParentsFromHybridshapes(shape, tempPart, geometry);
                        geometry.ParentGeometries = new List<ParentGeometry>();
                        if (inputList != null)
                        {
                            foreach (AnyObject hybrid in inputList)
                            {
                                ParentGeometry parentGeometry = new ParentGeometry();
                                if (hybrid != null)
                                {
                                    parentGeometry.Name = hybrid.get_Name();
                                }
                                else
                                {
                                    parentGeometry.Name = "";
                                }
                                geometry.ParentGeometries.Add(parentGeometry);
                            }
                        }
                        geometry.ShapeSpecific = new List<ShapeValues>();
                        if (shape is HybridShapePointCoord pointCoordinate)
                        {
                            ShapeValues pointshapeValues = new ShapeValues
                            {
                                XValue = pointCoordinate.X.Value,
                                YValue = pointCoordinate.Y.Value,
                                ZValue = pointCoordinate.Z.Value
                            };
                            (pointshapeValues.XMeasured, pointshapeValues.YMeasured, pointshapeValues.ZMeasured) = MethodsReadingGeometry.MeasuringPointCoordinate(tempPart, shape, Catia);
                            geometry.ShapeSpecific.Add(pointshapeValues);
                        }
                        else if (shape is HybridShapePointTangent) //for external reference points
                        {
                            ShapeValues shapeValues = new ShapeValues();
                            (shapeValues.XMeasured, shapeValues.YMeasured, shapeValues.ZMeasured) = MethodsReadingGeometry.MeasuringPointCoordinate(tempPart, shape, Catia);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        else if (shape is HybridShapeLinePtDir linePtDir)
                        {
                            ShapeValues shapeValues = new ShapeValues
                            {
                                BeginOffset = linePtDir.BeginOffset.Value,
                                EndOffset = linePtDir.EndOffset.Value
                            };
                            CatiaLengthType = linePtDir.GetLengthType();
                            shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        else if (shape is HybridShapeLinePtPt linePtPt)
                        {
                            ShapeValues shapeValues = new ShapeValues
                            {
                                BeginOffset = linePtPt.BeginOffset.Value,
                                EndOffset = linePtPt.EndOffset.Value
                            };
                            CatiaLengthType = linePtPt.GetLengthType();
                            shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        else if (shape is HybridShapeLineAngle lineAngle)
                        {
                            ShapeValues shapeValues = new ShapeValues
                            {
                                BeginOffset = lineAngle.BeginOffset.Value,
                                EndOffset = lineAngle.EndOffset.Value
                            };
                            CatiaLengthType = lineAngle.GetLengthType();
                            shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        else if (shape is HybridShapeLineNormal lineNormal)
                        {
                            ShapeValues shapeValues = new ShapeValues
                            {
                                BeginOffset = lineNormal.BeginOffset.Value,
                                EndOffset = lineNormal.EndOffset.Value
                            };
                            CatiaLengthType = lineNormal.GetLengthType();
                            shapeValues.LengthType = MethodsReadingGeometry.GetTheLengthTypeInString(CatiaLengthType);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        else if(geometry.Type =="") // WF Karosserie has external references without any specified type (but is a point)
                        {
                            geometry.Type = "HybridShapePointTangent";
                            ShapeValues shapeValues = new ShapeValues();
                            (shapeValues.XMeasured, shapeValues.YMeasured, shapeValues.ZMeasured) = MethodsReadingGeometry.MeasuringPointCoordinate(tempPart, shape, Catia);
                            geometry.ShapeSpecific.Add(shapeValues);
                        }
                        partClass.Geometries.Add(geometry);
                    }
                    MethodsReadingGeometry.AssignParentElementID(partClass);
                }
            }    
            return partClass;
        }
    }

}