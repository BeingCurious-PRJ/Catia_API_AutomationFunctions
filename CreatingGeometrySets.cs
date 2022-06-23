using HybridShapeTypeLib;
using MECMOD;
using INFITF;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;


namespace Automation-Functions_Methods
{
public class CreatingGeometrySets
    {        
        #region Fields
        private Application catiaApp;
        private Product kinematicProduct;
        public CreatingGeometrySets(Application catiaApp, Product kinematicProduct = null)
        {
            this.catiaApp = catiaApp;
            ProductDocument productDocument = CatiaHandler.GetActiveProductDocument(catiaApp);
            kinematicProduct = CatiaHandler.GetProductFromProductDocument(productDocument);
            this.kinematicProduct = kinematicProduct;
        }
        #endregion

        public MainGeometryClass CreateGeoSets(MainGeometryClass mainGeometry)
        {               
            foreach (GeoPartClass geoPartClass in mainGeometry.PartClass)
            {
                if (!string.IsNullOrEmpty(geoPartClass.PartName))
                {
                    Product currentProduct = CatiaHandler.GetProductByDefinition(kinematicProduct, geoPartClass.Definition);
                    Part currentPart = CatiaHandler.GetPartFromProduct(currentProduct);
                    bool doItAgain = false;
                    int iMax = 1000;
                    int i = 0;

                    do
                    {
                        i++;
                        doItAgain = false;
                        foreach (HybridBodyClass hybridBodyClass in geoPartClass.Hybridbodies)
                        {
                            if (hybridBodyClass.HybridbodyObject == null && hybridBodyClass.Name != "External References")
                            {
                                if (string.IsNullOrEmpty(hybridBodyClass.Parent))
                                {
                                    HybridBody hybridBody = MethodsCreatingGeometry.GetCatiaHybridbodyByHybridBodyClass(currentPart, hybridBodyClass);
                                    if(hybridBody != null)
                                    {
                                        hybridBodyClass.HybridbodyObject = hybridBody;
                                    }
                                    else
                                    {
                                        hybridBodyClass.HybridbodyObject = MethodsCreatingGeometry.CreateHybridBody(currentPart, hybridBodyClass.Name);                                 
                                    }
                                }
                                else
                                {
                                    HybridBody parentHB = MethodsCreatingGeometry.FindParentHybridbodyByID(hybridBodyClass.ParentGeoSetID, geoPartClass.Hybridbodies);
                                    if (parentHB != null)
                                    {                                  
                                        HybridBody hybridBody = MethodsCreatingGeometry.GetCatiaHybridbodyByHybridBodyClass(currentPart, hybridBodyClass);
                                        if (hybridBody != null)
                                        {
                                            hybridBodyClass.HybridbodyObject = hybridBody;
                                        }
                                        else
                                        {
                                            hybridBodyClass.HybridbodyObject = MethodsCreatingGeometry.CreateHybridBody(parentHB, hybridBodyClass.Name);
                                        }
                                     
                                    }
                                    else
                                    {
                                        doItAgain = true;
                                    }
                                }
                            }
                        }
                    } while (doItAgain || i < iMax);
                }
            }
            return mainGeometry;
        }

        public MainGeometryClass CreateGeoSetsElementsFromXml(MainGeometryClass mainGeometry, ExternalReferenceMatchClass matchClass)
        {
            foreach (GeoPartClass partClass in mainGeometry.PartClass)
            {
                if (!string.IsNullOrEmpty(partClass.PartName))
                {
                    Product currentProduct = CatiaHandler.GetProductByDefinition(kinematicProduct, partClass.Definition);
                    Part currentPart = CatiaHandler.GetPartFromProduct(currentProduct);
                    MethodsCreatingGeometry.GeometryBuilder(currentPart,kinematicProduct);

                    #region §Creating External References§
                    Selection selectedProductInstance = PointPublications.SelectionAndPasteSpecial(catiaApp, matchClass, partClass, currentPart,currentProduct);
                    GeoPartClass geoPartClass = MethodsCreatingGeometry.GetExternalPointsAsHybridShapeObjects(matchClass, partClass);
                    #endregion

                    bool doItAgain = false;
                    int iMax = 1000;
                    int i = 0;
                    do
                    {
                        i++;
                        foreach (Geometry geometry in geoPartClass.Geometries)
                        {
                            doItAgain = false;
                            if (geometry.HybridShapeObject == null && geometry.HybridBodyName != "External References")
                            {
                                if (geometry.ParentGeometries.Count != 0)
                                {
                                    foreach (ParentGeometry parent in geometry.ParentGeometries)
                                    {
                                        parent.HybridShapeObject = MethodsCreatingGeometry.GetParentElementCatiaObject(geoPartClass, parent);
                                        if (parent.HybridShapeObject == null)
                                        {
                                            doItAgain = true;
                                        }
                                    }
                                    if (!doItAgain)
                                    {
                                        CreateGeometry(geometry, currentPart, geoPartClass);
                                    }
                                }                               
                                else
                                {
                                    CreateGeometry(geometry, currentPart, geoPartClass);                                  
                                }
                            }
                        }
                    } while (doItAgain || i < iMax);
                }
            }
            return mainGeometry;
        }

        #region "Creating Visualization_GeoSet (Vis_Geo) Only"

        public MainGeometryClass CreateVisGeoSet(MainGeometryClass mainGeometry)
        {
            foreach (GeoPartClass geoPartClass in mainGeometry.PartClass)
            {
                if (!string.IsNullOrEmpty(geoPartClass.PartName))
                {
                    Product currentProduct = CatiaHandler.GetProductByDefinition(kinematicProduct, geoPartClass.Definition);
                    Part currentPart = CatiaHandler.GetPartFromProduct(currentProduct);

                    foreach (HybridBodyClass hybridBodyClass in geoPartClass.Hybridbodies)
                    {
                        if (hybridBodyClass.HybridbodyObject == null)
                        {
                            if (string.IsNullOrEmpty(hybridBodyClass.Parent) && hybridBodyClass.Name == "Vis Geo" || string.IsNullOrEmpty(hybridBodyClass.Parent) && hybridBodyClass.Name == "Vis Geom")
                            {
                                HybridBody hybridBody = MethodsCreatingGeometry.GetCatiaHybridbodyByHybridBodyClass(currentPart, hybridBodyClass);
                                if (hybridBody != null)
                                {
                                    hybridBodyClass.HybridbodyObject = hybridBody;
                                }
                                else
                                {
                                    hybridBodyClass.HybridbodyObject = MethodsCreatingGeometry.CreateHybridBody(currentPart, hybridBodyClass.Name);
                                }
                            }
                        }
                    }

                }
            }
            return mainGeometry;
        }
        public MainGeometryClass CreateVisGeoSetElementsFromXml(MainGeometryClass mainGeometry, ExternalReferenceMatchClass matchClass)
        {
            foreach (GeoPartClass partClass in mainGeometry.PartClass)
            {
                if (!string.IsNullOrEmpty(partClass.PartName))
                {
                    Product currentProduct = CatiaHandler.GetProductByDefinition(kinematicProduct, partClass.Definition);
                    Part currentPart = CatiaHandler.GetPartFromProduct(currentProduct);
                    MethodsCreatingGeometry.GeometryBuilder(currentPart, kinematicProduct);

                    #region §Creating External References§
                    Selection selectedProductInstance = PointPublications.SelectionAndPasteSpecial(catiaApp, matchClass, partClass, currentPart, currentProduct);
                    GeoPartClass geoPartClass = MethodsCreatingGeometry.GetExternalPointsAsHybridShapeObjects(matchClass, partClass);
                    #endregion

                    bool parentExists = true;
                    foreach (Geometry geometry in geoPartClass.Geometries)
                    {
                        if (geometry.HybridShapeObject == null && geometry.HybridBodyName == "Vis Geo" || geometry.HybridShapeObject == null && geometry.HybridBodyName == "Vis Geom")
                        {
                            if (geometry.ParentGeometries.Count != 0)
                            {
                                foreach (ParentGeometry parent in geometry.ParentGeometries)
                                {
                                    parent.HybridShapeObject = MethodsCreatingGeometry.GetParentElementCatiaObject(geoPartClass, parent);
                                    if (parent.HybridShapeObject == null)
                                    {
                                        parentExists = false;
                                    }
                                }
                                if (parentExists)
                                {
                                    CreateGeometry(geometry, currentPart, geoPartClass);
                                }
                            }
                        }
                    }                  
                }
            }
            return mainGeometry;
        }
        #endregion

        private void CreateGeometry(Geometry geometry, Part part, GeoPartClass geoPartClass)
        {
            if (!string.IsNullOrEmpty(geometry.Type))
            {
                string parentHybridBodyName = MethodsCreatingGeometry.GetParentHybridBodyName(geometry,geoPartClass);
                HybridBody hybridBody = MethodsCreatingGeometry.GetCatiaHybridbodyByGeometryInformation(part, geometry, parentHybridBodyName);
                HybridShape hybridShape = MethodsCreatingGeometry.GetCatiaElementByGeometryInformation(geometry,hybridBody);
                if(hybridShape != null)
                {
                    geometry.HybridShapeObject = hybridShape;
                }
                else if( hybridBody != null)
                {
                    switch (geometry.Type)
                    {
                        #region "Point"
                        case "HybridShapePointCoord":
                            PointCoordinateCreation(geometry, hybridBody);
                            break;
                        case "HybridShapePointBetween":
                            PointBetweenCreation(geometry, hybridBody);
                            break;
                        //case "HybridShapePointOnCurve": //use when vis geo elements in xml has complete parents
                        //    PointOnCurveCreation(geometry, hybridBody);
                        //    break;
                        #endregion

                        #region "Line"
                        case "HybridShapeLinePtPt":
                            LinePtPtCreation(geometry, hybridBody);
                            break;
                        case "HybridShapeLinePtDir":
                            LinePtDirCreation(geometry, hybridBody);
                            break;
                        case "HybridShapeLineAngle":
                            LineAngleCreation(geometry, hybridBody);
                            break;
                        case "HybridShapeLineNormal":
                            LineNormalCreation(geometry, hybridBody);
                            break;
                        #endregion

                        #region "Plane"
                        case "HybridShapePlaneNormal":
                            PlaneNormalCreation(geometry, hybridBody);
                            break;
                        case "HybridShapePlane3Points":
                            Plane3PointsCreation(geometry, hybridBody);
                            break;
                        case "HybridShapePlane1Line1Pt":
                            Plane1Lin1PtCreation(geometry, hybridBody);
                            break;
                        case "HybridShapePlane2Lines":
                            Plane2LinesCreation(geometry, hybridBody);
                            break;
                        case "HybridShapePlaneOffsetPt":
                            PlaneOffsetPointCreation(geometry, hybridBody);
                            break;
                        #endregion

                        #region "Circle"
                        case "HybridShapeCircleCtrRad":
                            CircleCtrPtAndRadiusCreation(geometry, hybridBody);
                            break;
                        case "HybridShapeCircleCtrPt":
                            CircleCtrAndPointCreation(geometry, hybridBody);
                            break;
                        #endregion

                        #region "Intersection, Projection, Symmetry"
                        case "HybridShapeIntersection":
                            IntersectionTwoObjects(geometry, hybridBody);
                            break;
                        case "HybridShapeProjection":
                            ProjectionOfTwoElements(geometry, hybridBody);
                            break;
                        case "HybridShapeSymmetry":
                            SymmetryWithElements(geometry, hybridBody);
                            break;
                        #endregion

                    }
                }
            }          
        }

        #region "Helper Methods for 'CreateGeometry' function"

        private static void SymmetryWithElements(Geometry geometry, HybridBody hybridBody)
        {
            HybridShape element = geometry.ParentGeometries[0].HybridShapeObject;
            HybridShape referenceObject = geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateSymmetryElementWithObjects(hybridBody, element, referenceObject, geometry.Name, true);
        }
        private static void ProjectionOfTwoElements(Geometry geometry, HybridBody hybridBody)
        {
            HybridShape object1 = geometry.ParentGeometries[0].HybridShapeObject;
            Plane object2 = (Plane)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateProjectionWithObjects(hybridBody, object1, object2, geometry.Name, true);
        }
        private static void IntersectionTwoObjects(Geometry geometry, HybridBody hybridBody)
        {
            HybridShape object1 = geometry.ParentGeometries[0].HybridShapeObject;
            HybridShape object2 = geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateIntersectionWithObjects(hybridBody, object1, object2, geometry.Name, true);
        }

        private static void CircleCtrAndPointCreation(Geometry geometry, HybridBody hybridBody)
        {
            Point point1 = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            Point point2 = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            Plane plane = (Plane)geometry.ParentGeometries[2].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateCircleWithCentreAndPoint(hybridBody, point1, point2, plane, geometry.Name, true, false);
        }
        private static void CircleCtrPtAndRadiusCreation(Geometry geometry, HybridBody hybridBody)
        {
            Point ctPoint = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            Plane plane = (Plane)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateCircleWithCentreAndRadius(hybridBody, ctPoint, plane, geometry.Name, 279, true, false);

        }

        private static void PlaneOffsetPointCreation(Geometry geometry, HybridBody hybridBody)
        {
            Plane plane = (Plane)geometry.ParentGeometries[0].HybridShapeObject;
            Point point = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePlaneFromOffsetPoint(hybridBody, plane, point, geometry.Name, true);
        }
        private static void Plane2LinesCreation(Geometry geometry, HybridBody hybridBody)
        {
            Line line1 = (Line)geometry.ParentGeometries[0].HybridShapeObject;
            Line line2 = (Line)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePlaneFrom2Lines(hybridBody, line1, line2, geometry.Name, true);
        }
        private static void Plane1Lin1PtCreation(Geometry geometry, HybridBody hybridBody)
        {
            Line line1P = (Line)geometry.ParentGeometries[0].HybridShapeObject;
            Point point1P = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePlaneFromLinePoint(hybridBody, line1P, point1P, geometry.Name, true);
        }
        private static void Plane3PointsCreation(Geometry geometry, HybridBody hybridBody)
        {
            Point pointP1 = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            Point pointP2 = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            Point pointP3 = (Point)geometry.ParentGeometries[2].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePlaneFrom3Points(hybridBody, pointP1, pointP2, pointP3, geometry.Name, true);
        }
        private static void PlaneNormalCreation(Geometry geometry, HybridBody hybridBody)
        {
            Line line = (Line)geometry.ParentGeometries[0].HybridShapeObject;
            Point pointP = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateNormalPlane(hybridBody, line, pointP, geometry.Name, true);
        }

        private static void LineNormalCreation(Geometry geometry, HybridBody hybridBody)
        {
            Plane plane = (Plane)geometry.ParentGeometries[0].HybridShapeObject;
            Point point = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFromPointNormal(hybridBody, plane, point, geometry.Name, true, false, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, geometry.ShapeSpecific[0].InternalLengthType, false);
        }
        private static void LineAngleCreation(Geometry geometry, HybridBody hybridBody)
        {
            Plane plane = (Plane)geometry.ParentGeometries[1].HybridShapeObject;
            Point point = (Point)geometry.ParentGeometries[2].HybridShapeObject;
            try
            {
                Line line = (Line)geometry.ParentGeometries[0].HybridShapeObject;
                geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFromAngle(hybridBody, line, plane, point, geometry.Name, true, false, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, 90, geometry.ShapeSpecific[0].InternalLengthType, false);
            }
            catch
            {
                if (geometry.HybridShapeObject == null)
                {
                    HybridShape projectionLine = geometry.ParentGeometries[0].HybridShapeObject;
                    geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFromAngleProjection(hybridBody, projectionLine, plane, point, geometry.Name, true, false, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, 90, geometry.ShapeSpecific[0].InternalLengthType, false);
                }
            }
        }
        private static void LinePtDirCreation(Geometry geometry, HybridBody hybridBody)
        {
            Point point = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            HybridShape directionElement = geometry.ParentGeometries[1].HybridShapeObject;
            Plane supportPlane = null;
            if(geometry.ParentGeometries.Count ==3)
            {
               supportPlane = (Plane)geometry.ParentGeometries[2].HybridShapeObject;
            }
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFromPointDirection(hybridBody, point, directionElement,supportPlane, geometry.Name, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, geometry.ShapeSpecific[0].InternalLengthType, true);
        }
        private static void LinePtPtCreation(Geometry geometry, HybridBody hybridBody)
        {
            HybridShape point1 = geometry.ParentGeometries[0].HybridShapeObject; ;
            HybridShape point2 = geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFrom2Points(hybridBody, point1, point2, geometry.Name, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, geometry.ShapeSpecific[0].InternalLengthType, true);

            //Point  point1 = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            //try
            //{
            //    Point point2 = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            //    geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFrom2Points(hybridBody, point1, point2, geometry.Name, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, geometry.ShapeSpecific[0].InternalLengthType, true);
            //}
            //catch
            //{
            //    HybridShape intersectionPoint = geometry.ParentGeometries[1].HybridShapeObject;
            //    //Plane supportPlane = (Plane)geometry.ParentGeometries[2].HybridShapeObject;
            //    geometry.HybridShapeObject = MethodsCreatingGeometry.CreateLineFromPtIntersection(hybridBody, point1, intersectionPoint, /*supportPlane,*/ geometry.Name, geometry.ShapeSpecific[0].BeginOffset, geometry.ShapeSpecific[0].EndOffset, geometry.ShapeSpecific[0].InternalLengthType, true);
            //}

        }

        //private static void PointOnCurveCreation(Geometry geometry, HybridBody hybridBody)
        //{
        //    Line linecurve = (Line)geometry.ParentGeometries[0].HybridShapeObject;
        //    Point point = (Point)geometry.ParentGeometries[1].HybridShapeObject;
        //    geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePointOnCurve(hybridBody, linecurve, point, geometry.Name, 0, true, false);
        //}

        private static void PointBetweenCreation(Geometry geometry, HybridBody hybridBody)
        {
            Point point1 = (Point)geometry.ParentGeometries[0].HybridShapeObject;
            Point point2 = (Point)geometry.ParentGeometries[1].HybridShapeObject;
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePointBetween(hybridBody, point1, point2, geometry.Name, 0.5, 0, true);
        }
        private static void PointCoordinateCreation(Geometry geometry, HybridBody hybridBody)
        {
            geometry.HybridShapeObject = MethodsCreatingGeometry.CreatePointCoordinate(hybridBody, geometry.Name, geometry.ShapeSpecific[0].XMeasured, geometry.ShapeSpecific[0].YMeasured, geometry.ShapeSpecific[0].ZMeasured, true);
        }
        #endregion
    }

}