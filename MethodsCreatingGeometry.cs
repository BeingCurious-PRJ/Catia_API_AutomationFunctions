using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;


namespace Automation-Functions_Methods
{
public class MethodsCreatingGeometry
    {
        #region "Fields"
        private static HybridShapeFactory HSF;
        public static  Part partobject;
        public static Product mainProduct;
        #endregion

        public static void GeometryBuilder(Part partObject, Product partProduct)
        {
            HSF = (HybridShapeFactory)partObject.HybridShapeFactory;
            partobject = partObject;
            mainProduct = partProduct;
        }
        public static void UpdatePart(AnyObject Element = null)
        {
            if (Element == null)
            {
                try
                {
                    partobject.Update();
                }
                catch (Exception)
                {

                }
            }
            else
            {
                try
                {
                    partobject.UpdateObject(Element);
                }
                catch (Exception)
                {
                }
            }
        }

        public static GeoPartClass GetExternalPointsAsHybridShapeObjects(ExternalReferenceMatchClass matchClass, GeoPartClass partClass)
        {
            foreach (PartElement partElement in matchClass.PartElements)
            {
                if (!string.IsNullOrEmpty(partElement.Name))
                {
                    if (partElement.Name == partClass.Definition)
                    {
                        foreach (ExternalReferences externalReferences in partElement.ExternalReference)
                        {
                            MatchAndAssignHybridShapeObject(partClass, externalReferences);
                        }
                    }
                }
            }

            return partClass;
        }
        private static void MatchAndAssignHybridShapeObject(GeoPartClass partClass, ExternalReferences externalReferences)
        {
            foreach (HybridBody hybridBody in partobject.HybridBodies)
            {
                if (hybridBody.get_Name() == "External References")
                {
                    foreach (HybridShape hybridShape in hybridBody.HybridShapes)
                    {
                        if (hybridShape.get_Name() == externalReferences.MatchPoint || hybridShape.get_Name() == externalReferences.Point)
                        {
                            foreach (Geometry geometry in partClass.Geometries)
                            {
                                if (geometry.Name == externalReferences.MatchPoint || geometry.Name == externalReferences.Point) //add a condition for x,y,z coordinate check too
                                {
                                    geometry.HybridShapeObject = hybridShape;
                                }
                            }
                        }
                    }
                }
            }
        }

        public static HybridBody CreateHybridBody(HybridBody targetHybridBody, string newName)
        {
            HybridBody resHybridBody = targetHybridBody.HybridBodies.Add();

            if (!string.IsNullOrEmpty(newName))
            {
                resHybridBody.set_Name(newName);
            }
            return resHybridBody;
        }
        public static HybridBody CreateHybridBody(Part part, string newName)
        {
            HybridBody resHybridBody = part.HybridBodies.Add();

            if (!string.IsNullOrEmpty(newName))
            {
                resHybridBody.set_Name(newName);
            }
            return resHybridBody;
        }
        public static HybridBody FindParentHybridbodyByID(string parentID, List<HybridBodyClass> hybridbodies)
        {
            foreach (HybridBodyClass hybridBodyClass in hybridbodies)
            {
                if (hybridBodyClass.GeoSetID == parentID)
                {
                    return hybridBodyClass.HybridbodyObject;
                }
            }
            return null;
        }
        public static HybridBody GetCatiaHybridbodyByGeometryInformation(Part part, Geometry geometry, string hbParentName)
        {
            foreach (HybridBody hb in part.HybridBodies)
            {
                if (geometry.HybridBodyName == hb.get_Name())
                {
                    return hb;
                }
                foreach (HybridBody subhb in hb.HybridBodies)
                {
                    if (geometry.HybridBodyName == subhb.get_Name() && hbParentName == hb.get_Name())
                    {
                        return subhb;
                    }
                    foreach (HybridBody sub1hb in subhb.HybridBodies )
                    {
                        if (geometry.HybridBodyName == sub1hb.get_Name() && hbParentName == subhb.get_Name())
                        {
                            return sub1hb;
                        }
                        foreach (HybridBody sub2hb in sub1hb.HybridBodies)
                        {
                            if (geometry.HybridBodyName == sub2hb.get_Name() && hbParentName == sub1hb.get_Name())
                            {
                                return sub2hb;
                            }
                        }
                    }
                }
            }
            return null;
        }
        public static string GetParentHybridBodyName(Geometry geometry, GeoPartClass geoPartClass)
        {
            foreach(HybridBodyClass hybridBody in geoPartClass.Hybridbodies)
            {
                if (geometry.GeoSetID == hybridBody.GeoSetID)
                {
                    return hybridBody.Parent;
                }
            }
            return null;
        }
        public static HybridBody GetCatiaHybridbodyByHybridBodyClass(Part part, HybridBodyClass hybridBody)
        {
            foreach (HybridBody hb in part.HybridBodies)
            {
                if (hybridBody.Name == hb.get_Name())
                {
                    return hb;
                }
                foreach (HybridBody subhb in hb.HybridBodies)
                {
                    if (hybridBody.Name == subhb.get_Name() && hybridBody.Parent == hb.get_Name())
                    {
                        return subhb;
                    }
                    foreach (HybridBody sub1hb in subhb.HybridBodies)
                    {
                        if (hybridBody.Name == sub1hb.get_Name() && hybridBody.Parent == subhb.get_Name())
                        {
                            return sub1hb;
                        }
                        foreach (HybridBody sub2hb in sub1hb.HybridBodies)
                        {
                            if (hybridBody.Name == sub2hb.get_Name() && hybridBody.Parent == sub1hb.get_Name())
                            {
                                return sub2hb;
                            }
                        }
                    }
                }
            }
            return null;
        }
        public static HybridShape GetCatiaElementByGeometryInformation(Geometry geometry,  HybridBody body)
        {
            if(body != null)
            {
                foreach (HybridShape shape in body.HybridShapes)
                {
                    if (geometry.Name == shape.get_Name())
                    {
                        return shape;
                    }
                }
            }
            return null;
        }
        public static HybridShape GetParentElementCatiaObject(GeoPartClass geoPartClass, ParentGeometry parent)
        {
            foreach (Geometry geometry in geoPartClass.Geometries)
            {
                if (parent.ElementID == geometry.ElementID)
                {
                    if (geometry.HybridShapeObject != null)
                    {
                        return geometry.HybridShapeObject;
                    }
                }
            }
            return null;
        }
        public static Reference GetReferenceInPartFromHybridshape(HybridShape hybridShape)
        {
            Reference result = null;
            //try
            //{
                result = partobject.CreateReferenceFromObject(hybridShape);
            //}
            //catch
            //{
            //    foreach (Product product in mainProduct.Products)
            //    {
            //        if (product.get_Definition() == "WF SKELETON" || product.get_Definition() == "SKELETON")
            //        {
            //            Part skeletonPart = CatiaHandler.GetPartFromProduct(product);
            //            result = skeletonPart.CreateReferenceFromObject(hybridShape);
            //        }
            //    }
               
            //}
            //include skeleton partobject to get the hybridshape object for external referencew
            return result;
        }
    
        #region POINT
        public static HybridShapePointCoord CreatePointCoordinate(HybridBody targetHybridBody, string pointName, double xVal = 0, double yVal = 0, double zVal = 0, bool updateAfterCreation = false)
        {
            HybridShapePointCoord newPoint;
            try
            {
                newPoint = HSF.AddNewPointCoord(xVal, yVal, zVal);
                if (!string.IsNullOrEmpty(pointName))
                {
                    newPoint.set_Name(ref pointName);
                }
                targetHybridBody.AppendHybridShape(newPoint);

                if (updateAfterCreation)
                {
                    UpdatePart(newPoint);
                }
            }
            catch
            {
                newPoint = null;
            }

            return newPoint;
        }
        public static HybridShapePointBetween CreatePointBetween(HybridBody targetHybridBody, Point iPoint1, Point iPoint2, string pointName, double iRatio = 1, int iOrientataion = 0, bool updateAfterCreation = false)
        {
            HybridShapePointBetween point = null;
            Reference initPointReference = GetReferenceInPartFromHybridshape(iPoint1);
            Reference finalPointReference = GetReferenceInPartFromHybridshape(iPoint2);
            try
            {
                point = HSF.AddNewPointBetween(initPointReference, finalPointReference,iRatio,iOrientataion);
                if (!string.IsNullOrEmpty(pointName))
                {
                    point.set_Name(ref pointName);
                }
                targetHybridBody.AppendHybridShape(point);

                if (updateAfterCreation)
                {
                    UpdatePart(point);
                }
            }
            catch
            {
                point = null;
            }
            return point;
        }
        public static HybridShapePointTangent CreatePointTangent(HybridBody targetHybridBody, Line iCurve, HybridShape iDirectionElement, string pointName, bool updateAfterCreation = false)
        {
            HybridShapePointTangent newPoint =null;
            Reference curveReference = GetReferenceInPartFromHybridshape(iCurve);
            Part part = MethodsReadingGeometry.GetParentPartObjectFromAnyObject(iDirectionElement);
            HybridShapeFactory shapeFactory = (HybridShapeFactory)part.HybridShapeFactory;
            HybridShapeDirection direction = shapeFactory.AddNewDirection(GetReferenceInPartFromHybridshape(iDirectionElement));

            try
            {
                newPoint = HSF.AddNewPointTangent(curveReference, direction);
                if (!string.IsNullOrEmpty(pointName))
                {
                    newPoint.set_Name(ref pointName);
                }
                targetHybridBody.AppendHybridShape(newPoint);

                if (updateAfterCreation)
                {
                    UpdatePart(newPoint);
                }
            }
            catch
            {
                newPoint = null;
            }

            return newPoint;
        }
        public static HybridShapePointOnCurve CreatePointOnCurve(HybridBody targetHybridBody, Line iCurve, Point iPoint, string pointName, double iLong = 0, bool updateAfterCreation = false, bool iOrientation = false)
        {
            HybridShapePointOnCurve newPoint =null;
            Reference lineReference = GetReferenceInPartFromHybridshape(iCurve);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                newPoint = HSF.AddNewPointOnCurveWithReferenceFromDistance(lineReference, pointReference, iLong, iOrientation);
                if (!string.IsNullOrEmpty(pointName))
                {
                    newPoint.set_Name(ref pointName);
                }
                targetHybridBody.AppendHybridShape(newPoint);

                if (updateAfterCreation)
                {
                    UpdatePart(newPoint);
                }
            }
            catch
            {
                newPoint = null;
            }

            return newPoint;
        }
        #endregion

        #region LINE
       
        public static HybridShapeLinePtPt CreateLineFrom2Points(HybridBody targetHybridBody,  HybridShape initPoint, HybridShape finalPoint,string lineName, double iBeginOffset = 0, double iEndOffset = 0, int lengthType =0,bool updateAfterCreation = false)
        {
            HybridShapeLinePtPt line = null;
            Reference initPointReference = GetReferenceInPartFromHybridshape(initPoint);
            Reference finalPointReference = GetReferenceInPartFromHybridshape(finalPoint);
            try
            {
                line = HSF.AddNewLinePtPt(initPointReference, finalPointReference);
                if (!string.IsNullOrEmpty(lineName))
                {
                    line.set_Name(ref lineName);
                }
                line.SetLengthType(lengthType);
                targetHybridBody.AppendHybridShape(line);

                if (updateAfterCreation)
                {
                    UpdatePart(line);
                }
            }
            catch
            {
                line = null;
            }
            return line;

        }
        //public static HybridShapeLinePtPt CreateLineFromPtIntersection(HybridBody targetHybridBody, Point initPoint, HybridShape intersectionPt,string lineName, double iBeginOffset = 0, double iEndOffset = 0, int lengthType = 0, bool updateAfterCreation = false)
        //{
        //    HybridShapeLinePtPt line = null;
        //    Reference initPointReference = GetReferenceInPartFromHybridshape(initPoint);
        //    Reference intersectionPtReference = GetReferenceInPartFromHybridshape(intersectionPt);
        //    //Reference planeReference = GetReferenceInPartFromHybridshape(supportPlane);
        //    try
        //    {
        //        line = HSF.AddNewLinePtPt(initPointReference, intersectionPtReference);
        //        if (!string.IsNullOrEmpty(lineName))
        //        {
        //            line.set_Name(ref lineName);
        //        }
        //        line.SetLengthType(lengthType);
        //        targetHybridBody.AppendHybridShape(line);

        //        if (updateAfterCreation)
        //        {
        //            UpdatePart(line);
        //        }
        //    }
        //    catch
        //    {
        //        line = null;
        //    }
        //    return line;
        //}
        public static HybridShapeLinePtDir CreateLineFromPointDirection(HybridBody targetHybridBody, Point iPoint, HybridShape iDirectionElement, Plane supportPlane, string line_Name, double iBeginOffset = 0, double iEndOffset = 0, int lengthType = 0, bool updateAfterCreation = false,bool iOrientation = false)
        {
           
            HybridShapeLinePtDir line = null;
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            HybridShapeDirection direction = HSF.AddNewDirection(GetReferenceInPartFromHybridshape(iDirectionElement));
            Reference planeReference = null;
            if (supportPlane != null)
            {
                planeReference = GetReferenceInPartFromHybridshape(supportPlane);
            }
            try
            {
                line = HSF.AddNewLinePtDir(pointReference, direction, iBeginOffset, iEndOffset, iOrientation);
                if (!string.IsNullOrEmpty(line_Name))
                {
                    line.set_Name(ref line_Name);
                }
                if(planeReference!= null)
                {
                    line.Support = planeReference;
                }
                line.SetLengthType(lengthType);
                targetHybridBody.AppendHybridShape(line);

                if (updateAfterCreation)
                {
                    UpdatePart(line);
                }
            }
            catch
            {
                line = null;
            }

            return line;
        }       
        public static HybridShapeLineAngle CreateLineFromAngle(HybridBody targetHybridBody, Line iCurve, Plane iSurface, Point iPoint , string lineName ,bool updateAfterCreation = false, bool iGeodesic = false, double iBeginOffset = 0, double iEndOffset = 0, double iAngle = 0, int lengthType = 0, bool iOrientation = false)
        {
            HybridShapeLineAngle line =null;
            Reference lineReference = GetReferenceInPartFromHybridshape(iCurve);
            Reference surfaceReference = GetReferenceInPartFromHybridshape(iSurface);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                line = HSF.AddNewLineAngle(lineReference, surfaceReference, pointReference, iGeodesic, iBeginOffset, iEndOffset, iAngle, iOrientation);
                if (!string.IsNullOrEmpty(lineName))
                {
                    line.set_Name(ref lineName);
                }
                line.SetLengthType(lengthType);
                targetHybridBody.AppendHybridShape(line);

                if (updateAfterCreation)
                {
                    UpdatePart(line);
                }
            }
            catch
            {
                line = null;
            }

            return line;
        }
        public static HybridShapeLineAngle CreateLineFromAngleProjection(HybridBody targetHybridBody, HybridShape iCurve, Plane iSurface, Point iPoint, string lineName, bool updateAfterCreation = false, bool iGeodesic = false, double iBeginOffset = 0, double iEndOffset = 0, double iAngle = 0, int lengthType = 0, bool iOrientation = false)
        {
            HybridShapeLineAngle line = null;
            Reference projectedlineReference = GetReferenceInPartFromHybridshape(iCurve);
            Reference surfaceReference = GetReferenceInPartFromHybridshape(iSurface);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                line = HSF.AddNewLineAngle(projectedlineReference, surfaceReference, pointReference, iGeodesic, iBeginOffset, iEndOffset, iAngle, iOrientation);
                if (!string.IsNullOrEmpty(lineName))
                {
                    line.set_Name(ref lineName);
                }
                line.SetLengthType(lengthType);
                targetHybridBody.AppendHybridShape(line);

                if (updateAfterCreation)
                {
                    UpdatePart(line);
                }
            }
            catch
            {
                line = null;
            }

            return line;
        }
        public static HybridShapeLineNormal CreateLineFromPointNormal(HybridBody targetHybridBody, Plane iSurface, Point iPoint, string lineName, bool updateAfterCreation = false, bool iGeodesic = false, double iBeginOffset = 0, double iEndOffset = 0, int lengthType = 0, bool iOrientation = false)
        {
            HybridShapeLineNormal line = null;
            Reference surfaceReference = GetReferenceInPartFromHybridshape(iSurface);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                line = HSF.AddNewLineNormal(surfaceReference, pointReference,iBeginOffset,iEndOffset,iOrientation);
                if (!string.IsNullOrEmpty(lineName))
                {
                    line.set_Name(ref lineName);
                }
                line.SetLengthType(lengthType);
                targetHybridBody.AppendHybridShape(line);

                if (updateAfterCreation)
                {
                    UpdatePart(line);
                }
            }
            catch
            {
                line = null;
            }
            return line;
        }
        #endregion

        #region PLANE
        public static HybridShapePlaneNormal CreateNormalPlane(HybridBody targetHybridBody,Line iLine , Point iPoint, string planeName, bool updateAfterCreation = false)
        {
            HybridShapePlaneNormal plane = null;
            Reference lineReference = GetReferenceInPartFromHybridshape(iLine);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                plane = HSF.AddNewPlaneNormal(lineReference, pointReference);
                if (!string.IsNullOrEmpty( planeName))
                {
                    plane.set_Name(planeName);
                }
                targetHybridBody.AppendHybridShape(plane);

                if (updateAfterCreation)
                {
                    UpdatePart(plane);

                }
            }
            catch
            {
                plane = null;
            }
            return plane;
        }
        public static HybridShapePlane1Line1Pt CreatePlaneFromLinePoint(HybridBody targetHybridBody, Line iLine , Point iPoint, string planeName, bool updateAfterCreation = false)
        {
            HybridShapePlane1Line1Pt plane = null;
            Reference lineReference = GetReferenceInPartFromHybridshape(iLine);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                plane = HSF.AddNewPlane1Line1Pt(lineReference, pointReference);
                if (!string.IsNullOrEmpty(planeName))
                {
                    plane.set_Name(planeName);
                }
                targetHybridBody.AppendHybridShape(plane);

                if (updateAfterCreation)
                {
                    UpdatePart(plane);
                }
            }
            catch
            {
                plane = null;
            }

            return plane;
        }
        public static HybridShapePlane3Points CreatePlaneFrom3Points(HybridBody targetHybridBody,Point iPoint1, Point iPoint2, Point iPoint3, string planeName, bool updateAfterCreation = false)
        {
            HybridShapePlane3Points plane =null;
            Reference point1Reference = GetReferenceInPartFromHybridshape(iPoint1);
            Reference point2Reference = GetReferenceInPartFromHybridshape(iPoint2);
            Reference point3Reference = GetReferenceInPartFromHybridshape(iPoint3);

            try
            {
                plane = HSF.AddNewPlane3Points(point1Reference, point2Reference, point3Reference);
                if (!string.IsNullOrEmpty( planeName))
                {
                    plane.set_Name(planeName);
                }
                targetHybridBody.AppendHybridShape(plane);

                if (updateAfterCreation)
                {
                    UpdatePart(plane);

                }
            }
            catch
            {
                plane = null;
            }

            return plane;
        }
        public static HybridShapePlane2Lines CreatePlaneFrom2Lines(HybridBody targetHybridBody,Line iLine1, Line iLine2, string planeName, bool updateAfterCreation = false)
        {
            HybridShapePlane2Lines plane = null;
            Reference line1Reference = GetReferenceInPartFromHybridshape(iLine1);
            Reference line2Reference = GetReferenceInPartFromHybridshape(iLine2);
            try
            {
                plane = HSF.AddNewPlane2Lines(line1Reference, line2Reference);
                if (!string.IsNullOrEmpty(planeName))
                {
                    plane.set_Name(planeName);
                }
                targetHybridBody.AppendHybridShape(plane);

                if (updateAfterCreation)
                {
                    UpdatePart(plane);
                }
            }
            catch
            {
                plane = null;
            }

            return plane;
        }
        public static HybridShapePlaneOffsetPt CreatePlaneFromOffsetPoint(HybridBody targetHybridBody, Plane iPlane, Point iPoint, string planeName, bool updateAfterCreation = false)
        {
            HybridShapePlaneOffsetPt plane = null;
            Reference planeReference = GetReferenceInPartFromHybridshape(iPlane);
            Reference pointReference = GetReferenceInPartFromHybridshape(iPoint);
            try
            {
                plane = HSF.AddNewPlaneOffsetPt(planeReference, pointReference);
                if (!string.IsNullOrEmpty(planeName))
                {
                    plane.set_Name(planeName);
                }
                targetHybridBody.AppendHybridShape(plane);

                if (updateAfterCreation)
                {
                    UpdatePart(plane);
                }
            }
            catch
            {
                plane = null;
            }

            return plane;

        }
        #endregion

        #region CIRCLE
        public static HybridShapeCircleCtrRad CreateCircleWithCentreAndRadius(HybridBody targetHybridBody, Point cenPoint, Plane supportPlane, string circleName, double radius = 0,bool updateAfterCreation = false, bool iGeodesic = false)
        {
            HybridShapeCircleCtrRad circle =null;
            Reference planeReference = GetReferenceInPartFromHybridshape(supportPlane);
            Reference pointReference = GetReferenceInPartFromHybridshape(cenPoint);

            try
            {
                circle = HSF.AddNewCircleCtrRad(pointReference, planeReference, iGeodesic, radius);
                if (!string.IsNullOrEmpty(circleName))
                {
                    circle.set_Name(circleName);
                }
                targetHybridBody.AppendHybridShape(circle);
                if (updateAfterCreation)
                {
                    UpdatePart(circle);

                }
            }
            catch
            {
                circle = null;
            }

            return circle;
        }
        public static HybridShapeCircleCtrPt CreateCircleWithCentreAndPoint (HybridBody targetHybridBody, Point cenPoint, Point crossingPoint, Plane supportPlane, string circleName, bool updateAfterCreation = false, bool iGeodesic = false)
        {
            HybridShapeCircleCtrPt circle = null;          
            Reference point1Reference = GetReferenceInPartFromHybridshape(cenPoint);
            Reference point2Reference = GetReferenceInPartFromHybridshape(crossingPoint);
            Reference planeReference = GetReferenceInPartFromHybridshape(supportPlane);

            try
            {
                circle = HSF.AddNewCircleCtrPt(point1Reference, point2Reference, planeReference, iGeodesic);
                if (!string.IsNullOrEmpty(circleName))
                {
                    circle.set_Name(circleName);
                }
                targetHybridBody.AppendHybridShape(circle);
                if (updateAfterCreation)
                {
                    UpdatePart(circle);

                }
            }
            catch
            {
                circle = null;
            }

            return circle;

        }
        #endregion

        #region INTERSECTION, PROJECTION, SYMMETRY
        public static HybridShapeIntersection CreateIntersectionWithObjects(HybridBody targetHybridBody, HybridShape iObject1, HybridShape iObject2, string intersectionName, bool updateAfterCreation = false)
        {
            HybridShapeIntersection intersection = null;
            Reference refObject1 = GetReferenceInPartFromHybridshape(iObject1);
            Reference refObject2 = GetReferenceInPartFromHybridshape(iObject2);
            try
            {
                intersection = HSF.AddNewIntersection(refObject1, refObject2);
                if(!string.IsNullOrEmpty(intersectionName))
                {
                    intersection.set_Name(intersectionName);
                }
                targetHybridBody.AppendHybridShape(intersection);
                if(updateAfterCreation)
                {
                    UpdatePart(intersection);
                }
            }
            catch
            {
                intersection = null;
            }
            return intersection;
        }
        public static HybridShapeProject CreateProjectionWithObjects(HybridBody targetHybridBody, HybridShape iObject1, Plane iObject2, string projectionName, bool updateAfterCreation = false)
        {
            HybridShapeProject shapeProject = null;
            Reference iElement1Ref  = GetReferenceInPartFromHybridshape(iObject1);
            Reference iSupportRef2 = GetReferenceInPartFromHybridshape(iObject2);
            try
            {
                shapeProject = HSF.AddNewProject(iElement1Ref, iSupportRef2);
                if (!string.IsNullOrEmpty(projectionName))
                {
                    shapeProject.set_Name(projectionName);
                }
                targetHybridBody.AppendHybridShape(shapeProject);
                if (updateAfterCreation)
                {
                    UpdatePart(shapeProject);
                }
            }
            catch
            {
                shapeProject = null;
            }
            return shapeProject;
           
        }
        public static HybridShapeSymmetry CreateSymmetryElementWithObjects(HybridBody targetHybridBody, HybridShape iElement, HybridShape iReference, string symmetryName, bool updateAfterCreation = false)
        {
            HybridShapeSymmetry symmetry = null;
            Reference iElementRef = GetReferenceInPartFromHybridshape(iElement);
            Reference iRef2 = GetReferenceInPartFromHybridshape(iReference);
            try
            {
                symmetry = HSF.AddNewSymmetry(iElementRef, iRef2);
                if (!string.IsNullOrEmpty(symmetryName))
                {
                    symmetry.set_Name(symmetryName);
                }
                targetHybridBody.AppendHybridShape(symmetry);
                if (updateAfterCreation)
                {
                    UpdatePart(symmetry);
                }
            }
            catch
            {
                symmetry = null;
            }
            return symmetry;


        }
        #endregion

        
    }

}