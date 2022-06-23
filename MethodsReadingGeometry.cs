using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using SPATypeLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace Automation-Functions_Methods
{
    public class MethodsReadingGeometry
    {
        public static void AssignParentElementID(GeoPartClass partClass)
        {
            foreach (Geometry geometryElement in partClass.Geometries)
            {
                foreach (ParentGeometry parentGeometry in geometryElement.ParentGeometries)
                {
                    foreach (Geometry geometry in partClass.Geometries)
                    {
                        if (parentGeometry.Name == geometry.Name)
                        {
                            parentGeometry.ElementID = geometry.ElementID;
                        }
                    }
                }
            }
        }
        public static (double x, double y, double z) MeasuringPointCoordinate(Part tempPart, HybridShape shape, INFITF.Application Catia)
        {
            SPAWorkbench spaWorkbench = (SPAWorkbench)Catia.ActiveDocument.GetWorkbench("SPAWorkbench");//interface to the catia "space analysis workbench"
            Reference pointReference = tempPart.CreateReferenceFromObject(shape);
            Measurable measurable = spaWorkbench.GetMeasurable(pointReference);
            object[] coord = new object[3];
            double x, y, z;
            try
            {
                measurable.GetPoint(coord);
                x = (double)coord[0];
                y = (double)coord[1];
                z = (double)coord[2];
            }
            catch
            {
                x = 0;
                y = 0;
                z = 0;
            }
            return (x, y, z);
        }

        #region "Self-Loop functions getting parent Part,GeoSet

        public static Part GetParentPartObjectFromAnyObject(AnyObject anyObject)
        {
            if (anyObject.Parent is Part part)
            {
                return part;
            }
            else
            {
                return GetParentPartObjectFromAnyObject((AnyObject)anyObject.Parent);
            }
        }
        public static AnyObject GetParentGeoSetFromChildGeoSets(AnyObject anyObject)
        {
            if (anyObject.Parent is Part)
            {
                return anyObject;
            }
            else
            {
                return GetParentGeoSetFromChildGeoSets((AnyObject)anyObject.Parent);
            }
        }
        #endregion

        #region "Getting HBs of all parts in product"

        public static List<HybridBody> GetAllHybridbodiesFromParts(List<Part> partList)
        {
            List<HybridBody> hybridBodies = null;
            foreach (Part item in partList)
            {
                hybridBodies = GetHybridBodies(item, hybridBodies);
            }
            return hybridBodies;
        }
        public static List<HybridBody> GetHybridBodies(AnyObject catiaObject, List<HybridBody> hybridBodies = null)
        {
            if (hybridBodies == null)
            {
                hybridBodies = new List<HybridBody>();
            }
            foreach (HybridBody item in GetAllHybridbodies(catiaObject))
            {
                if (item != null)
                {
                    hybridBodies.Add(item);
                    hybridBodies = GetHybridBodies(item, hybridBodies);
                }
                else
                {
                    MessageBox.Show("GetAllHybridbodies returned null");
                }
            }
            return hybridBodies;
        }
        public static HybridBodies GetAllHybridbodies(AnyObject anyObject)
        {
            if (anyObject is Body body)
            {
                return body.HybridBodies;
            }
            else if (anyObject is MECMOD.HybridBody hybridBody)
            {
                return hybridBody.HybridBodies;
            }
            else if (anyObject is Part part)
            {
                return part.HybridBodies;
            }
            return null;
        }
        #endregion

        #region "Methods to get parent references for each Hybrid Shape element"

        public static List<AnyObject> GetInputParentsFromHybridshapes(HybridShape hybridShape, Part part, Geometry geometry)
        {
            HybridShapeFactory hybridShapeFactory = (HybridShapeFactory)part.HybridShapeFactory;
            List<AnyObject> resultList = new List<AnyObject>();
            _ = new List<Reference>();
            List<Reference> references;
            switch (hybridShape)
            {
                #region $ POINT $
                case HybridShapePointCoord pointcoord:
                    geometry.Type = "HybridShapePointCoord";
                    references = GetInputReferencesFromPointCoord(pointcoord);
                    break;
                case HybridShapePointBetween pointBetween:
                    geometry.Type = "HybridShapePointBetween";
                    references = GetInputReferencesFromPointBetween(pointBetween);
                    break;
                case HybridShapePointOnCurve pointOnCurve:
                    geometry.Type = "HybridShapePointOnCurve";
                    references = GetInputReferencesFromPointOnCurve(pointOnCurve);
                    break;
                case HybridShapePointOnPlane pointOnPlane:
                    geometry.Type = "HybridShapePointOnPlane";
                    references = GetInputReferencesFromPointOnPlane(pointOnPlane);
                    break;
                case HybridShapePointOnSurface pointOnSurface:
                    geometry.Type = "HybridShapePointOnSurface";
                    references = GetInputReferencesFromPointOnSurface(pointOnSurface);
                    break;
                case HybridShapePointCenter pointCentre:
                    geometry.Type = "HybridShapePointCenter";
                    references = GetInputReferencesFromPointCentre(pointCentre);
                    break;
                case HybridShapePointTangent pointTangent:
                    geometry.Type = "HybridShapePointTangent ";
                    references = GetInputReferencesFromPointTangent(pointTangent);
                    break;
                #endregion

                #region $ LINE $
                case HybridShapeLinePtPt lineptpt:
                    geometry.Type = "HybridShapeLinePtPt";
                    references = GetInputReferencesFromLinePtPt(lineptpt);
                    break;
                case HybridShapeLinePtDir linePtDir:
                    geometry.Type = "HybridShapeLinePtDir";
                    references = GetInputReferencesFromLinePtDir(linePtDir, part);
                    break;
                case HybridShapeLineAngle lineAngle:
                    geometry.Type = "HybridShapeLineAngle";
                    references = GetInputReferencesFromLineAngle(lineAngle);
                    break;
                case HybridShapeLineNormal lineNormal:
                    geometry.Type = "HybridShapeLineNormal";
                    references = GetInputReferencesFromLineNormal(lineNormal);
                    break;
                case HybridShapeLineTangency lineTangency:
                    geometry.Type = "HybridShapeLineTangency";
                    references = GetInputReferencesFromLineTangency(lineTangency);
                    break;
                case HybridShapeLineBisecting lineBisecting:
                    geometry.Type = "HybridShapeLineBisecting";
                    references = GetInputReferencesFromLineBisecting(lineBisecting);
                    break;
                #endregion

                #region $ PLANE $
                case HybridShapePlaneNormal planenormal:
                    geometry.Type = "HybridShapePlaneNormal";
                    references = GetInputReferencesFromPlaneNormal(planenormal);
                    break;
                case HybridShapePlane1Line1Pt plane1Line1Pt:
                    geometry.Type = "HybridShapePlane1Line1Pt";
                    references = GetInputReferencesFromPlane1Line1Pt(plane1Line1Pt);
                    break;
                case HybridShapePlane2Lines plane2Lines:
                    geometry.Type = "HybridShapePlane2Lines";
                    references = GetInputReferencesFromPlane2Lines(plane2Lines);
                    break;
                case HybridShapePlane3Points plane3Points:
                    geometry.Type = "HybridShapePlane3Points";
                    references = GetInputReferencesFromPlane3Points(plane3Points);
                    break;
                case HybridShapePlaneAngle planeAngle:
                    geometry.Type = "HybridShapePlaneAngle";
                    references = GetInputReferencesFromPlaneAngle(planeAngle);
                    break;
                case HybridShapePlaneOffset planeOffset:
                    geometry.Type = "HybridShapePlaneOffset";
                    references = GetInputReferencesFromPlaneOffset(planeOffset);
                    break;
                case HybridShapePlaneOffsetPt planeOffsetPt:
                    geometry.Type = "HybridShapePlaneOffsetPt";
                    references = GetInputReferencesFromPlaneOffsetPt(planeOffsetPt);
                    break;
                case HybridShapePlaneTangent planeTangent:
                    geometry.Type = "HybridShapePlaneTangent";
                    references = GetInputReferencesFromPlaneTangent(planeTangent);
                    break;
                #endregion

                #region $ CIRCLE $
                case HybridShapeCircleCtrRad circleCtrRad:
                    geometry.Type = "HybridShapeCircleCtrRad";
                    references = GetInputReferencesFromCircleCtrRad(circleCtrRad);
                    break;
                case HybridShapeCircleCtrPt circleCtrPt:
                    geometry.Type = "HybridShapeCircleCtrPt";
                    references = GetInputReferencesFromCircleCtrPt(circleCtrPt);
                    break;
                #endregion

                #region $ INTERSECTION, PROJECTION, SYMMETRY $
                case HybridShapeIntersection shapeIntersection:
                    geometry.Type = "HybridShapeIntersection";
                    references = GetInputReferencesFromIntersectionObjects(shapeIntersection);
                    break;
                case HybridShapeProject shapeProject:
                    geometry.Type = "HybridShapeProjection";
                    references = GetInputReferencesFromProjectionObjects(shapeProject);
                    break;
                case HybridShapeSymmetry shapeSymmetry:
                    geometry.Type = "HybridShapeSymmetry";
                    references = GetInputReferencesFromSymmetryObjects(shapeSymmetry);
                    break;
                #endregion

                default:
                    references = null;
                    break;
            }
            #region null check and return hybridshape object from references
            if (references != null)
            {
                foreach (Reference item in references)
                {
                    try
                    {
                        if (item != null) //tried for skeleton
                        {
                            resultList.Add(hybridShapeFactory.GSMGetObjectFromReference(item));
                        }
                    }
                    catch (InvalidCastException)
                    {
                        resultList.Add(null);
                    }
                    catch (COMException)
                    {
                        resultList.Add(null);
                    }
                }
            }
            else
            {
                resultList = null;
            }
            #endregion

            return resultList;
        }

        #region "Point"
        public static List<Reference> GetInputReferencesFromPointCoord(HybridShapePointCoord pointcoord)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointcoord.PtRef != null)
                {
                    references.Add(pointcoord.PtRef);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (pointcoord.RefAxisSystem != null)
                {
                    references.Add(pointcoord.RefAxisSystem);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointBetween(HybridShapePointBetween pointBetween)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointBetween.FirstPoint != null)
                {
                    references.Add(pointBetween.FirstPoint);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (pointBetween.SecondPoint != null)
                {
                    references.Add(pointBetween.SecondPoint);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointOnCurve(HybridShapePointOnCurve pointOnCurve)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointOnCurve.Curve != null)
                {
                    references.Add(pointOnCurve.Curve);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (pointOnCurve.Point != null)
                {
                    references.Add(pointOnCurve.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointOnPlane(HybridShapePointOnPlane pointOnPlane)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointOnPlane.Plane != null)
                {
                    references.Add(pointOnPlane.Plane);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (pointOnPlane.Point != null)
                {
                    references.Add(pointOnPlane.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointOnSurface(HybridShapePointOnSurface pointOnSurface)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointOnSurface.Surface != null)
                {
                    references.Add(pointOnSurface.Surface);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (pointOnSurface.Point != null)
                {
                    references.Add(pointOnSurface.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointCentre(HybridShapePointCenter pointCentre)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointCentre.Element != null)
                {
                    references.Add(pointCentre.Element);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPointTangent(HybridShapePointTangent pointTangent)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (pointTangent.Curve != null)
                {
                    references.Add(pointTangent.Curve);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if ((Reference)pointTangent.Direction != null)
                {
                    references.Add((Reference)pointTangent.Direction);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        #endregion

        #region "Line"
        public static string GetTheLengthTypeInString(int lengthType)
        {
            string lengthTypeDetail = "";
            switch (lengthType)
            {
                case 0:
                    lengthTypeDetail = "Length";
                    break;
                case 1:
                    lengthTypeDetail = "Infinite";
                    break;
                case 2:
                    lengthTypeDetail = "InfiniteStartPoint";
                    break;
                case 3:
                    lengthTypeDetail = "InfiniteEndPoint";
                    break;
                default:
                    break;
            }
            return lengthTypeDetail;
        }
        public static List<Reference> GetInputReferencesFromLinePtPt(HybridShapeLinePtPt line)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (line.PtOrigine != null)
                {
                    references.Add(line.PtOrigine);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (line.PtExtremity != null)
                {
                    references.Add(line.PtExtremity);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (line.Support != null)
                {
                    references.Add(line.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromLinePtDir(HybridShapeLinePtDir linePtDir, Part part)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (linePtDir.Point != null)
                {
                    references.Add(linePtDir.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if ((Reference)linePtDir.Dir != null)
                {
                    references.Add(linePtDir.Dir.Object);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (linePtDir.Support != null)
                {
                    references.Add(linePtDir.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromLineAngle(HybridShapeLineAngle lineAngle)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (lineAngle.Curve != null)
                {
                    references.Add(lineAngle.Curve);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineAngle.Surface != null)
                {
                    references.Add(lineAngle.Surface);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineAngle.Point != null)
                {
                    references.Add(lineAngle.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromLineNormal(HybridShapeLineNormal lineNormal)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (lineNormal.Surface != null)
                {
                    references.Add(lineNormal.Surface);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineNormal.Point != null)
                {
                    references.Add(lineNormal.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromLineTangency(HybridShapeLineTangency lineTangency)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (lineTangency.Curve != null)
                {
                    references.Add(lineTangency.Curve);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineTangency.Point != null)
                {
                    references.Add(lineTangency.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineTangency.Support != null)
                {
                    references.Add(lineTangency.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromLineBisecting(HybridShapeLineBisecting lineBisecting)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (lineBisecting.Elem1 != null)
                {
                    references.Add(lineBisecting.Elem1);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineBisecting.Elem2 != null)
                {
                    references.Add(lineBisecting.Elem2);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineBisecting.RefPoint != null)
                {
                    references.Add(lineBisecting.RefPoint);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (lineBisecting.Support != null)
                {
                    references.Add(lineBisecting.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        #endregion

        #region "Plane"

        public static List<Reference> GetInputReferencesFromPlaneNormal(HybridShapePlaneNormal planenormal)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (planenormal.Curve != null)
                {
                    references.Add(planenormal.Curve);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (planenormal.Point != null)
                {
                    references.Add(planenormal.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlane1Line1Pt(HybridShapePlane1Line1Pt plane1Line1Pt)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (plane1Line1Pt.Line != null)
                {
                    references.Add(plane1Line1Pt.Line);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (plane1Line1Pt.Point != null)
                {
                    references.Add(plane1Line1Pt.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlane2Lines(HybridShapePlane2Lines plane2Lines)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (plane2Lines.First != null)
                {
                    references.Add(plane2Lines.First);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (plane2Lines.Second != null)
                {
                    references.Add(plane2Lines.Second);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlane3Points(HybridShapePlane3Points plane3Points)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (plane3Points.First != null)
                {
                    references.Add(plane3Points.First);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (plane3Points.Second != null)
                {
                    references.Add(plane3Points.Second);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (plane3Points.Third != null)
                {
                    references.Add(plane3Points.Third);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlaneAngle(HybridShapePlaneAngle planeAngle)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (planeAngle.Plane != null)
                {
                    references.Add(planeAngle.Plane);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (planeAngle.RevolAxis != null)
                {
                    references.Add(planeAngle.RevolAxis);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlaneOffset(HybridShapePlaneOffset planeOffset)
        {
            List<Reference> references = new List<Reference>
            {
                planeOffset.Plane
            };
            try
            {
                if (planeOffset.Plane != null)
                {
                    references.Add(planeOffset.Plane);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlaneOffsetPt(HybridShapePlaneOffsetPt planeOffsetPt)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (planeOffsetPt.Plane != null)
                {
                    references.Add(planeOffsetPt.Plane);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (planeOffsetPt.Point != null)
                {
                    references.Add(planeOffsetPt.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        public static List<Reference> GetInputReferencesFromPlaneTangent(HybridShapePlaneTangent planeTangent)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (planeTangent.Surface != null)
                {
                    references.Add(planeTangent.Surface);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (planeTangent.Point != null)
                {
                    references.Add(planeTangent.Point);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        #endregion

        #region "Circle"

        public static List<Reference> GetInputReferencesFromCircleCtrRad(HybridShapeCircleCtrRad circleCtrRad)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (circleCtrRad.Center != null)
                {
                    references.Add(circleCtrRad.Center);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (circleCtrRad.Support != null)
                {
                    references.Add(circleCtrRad.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            //try
            //{
            //    if (circleCtrRad.Support != null)
            //    {
            //        references.Add((Reference)circleCtrRad.Radius);
            //    }
            //}
            //catch
            //{
            //    references.Add(null);
            //}
            return references;
        }
        private static List<Reference> GetInputReferencesFromCircleCtrPt(HybridShapeCircleCtrPt circleCtrPt)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (circleCtrPt.Center != null)
                {
                    references.Add(circleCtrPt.Center);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (circleCtrPt.Center != null)
                {
                    references.Add(circleCtrPt.CrossingPoint);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (circleCtrPt.Support != null)
                {
                    references.Add(circleCtrPt.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        #endregion

        #region "Intersection, Projection, Symmetry"
        private static List<Reference> GetInputReferencesFromIntersectionObjects(HybridShapeIntersection shapeIntersection)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (shapeIntersection.Element1 != null)
                {
                    references.Add(shapeIntersection.Element1);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (shapeIntersection.Element2 != null)
                {
                    references.Add(shapeIntersection.Element2);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        private static List<Reference> GetInputReferencesFromProjectionObjects(HybridShapeProject shapeProject)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (shapeProject.ElemToProject != null)
                {
                    references.Add(shapeProject.ElemToProject);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (shapeProject.Support != null)
                {
                    references.Add(shapeProject.Support);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }
        private static List<Reference> GetInputReferencesFromSymmetryObjects(HybridShapeSymmetry shapeSymmetry)
        {
            List<Reference> references = new List<Reference>();
            try
            {
                if (shapeSymmetry.ElemToSymmetry!= null)
                {
                    references.Add(shapeSymmetry.ElemToSymmetry);
                }
            }
            catch
            {
                references.Add(null);
            }
            try
            {
                if (shapeSymmetry.Reference != null)
                {
                    references.Add(shapeSymmetry.Reference);
                }
            }
            catch
            {
                references.Add(null);
            }
            return references;
        }

        #endregion

        #endregion
    }

}