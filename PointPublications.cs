using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation-Functions_Methods
{
   public class PointPublications
    {
        //private static List<string> publicationInSkeleton;
        public static bool check = false;

        //public static List<string> PublicationInSkeleton { get => publicationInSkeleton; set => publicationInSkeleton = value; }

        public static Selection SelectionAndPasteSpecial(Application catiaApp,ExternalReferenceMatchClass matchClass, GeoPartClass partClass, Part currentPart,Product currentProduct)
        {
            ProductDocument activeProductDocument = CatiaHandler.GetActiveProductDocument(catiaApp);
            Product catiaProduct = CatiaHandler.GetProductFromProductDocument(activeProductDocument);
            Selection selectedCatiaProduct = activeProductDocument.Selection;

            foreach (Product product in catiaProduct.Products)
            {
                if(product.get_Definition() == "WF SKELETON" || product.get_Definition() == "SKELETON")
                {
                    foreach(PartElement partElement in matchClass.PartElements)
                    {
                        if(partClass.Definition == partElement.Name && partClass.Definition == currentProduct.get_Definition())
                        {
                            foreach (ExternalReferences external in partElement.ExternalReference)
                            {
                                external.Exists = false;
                                ExternalReferences externalRef = CheckIfExternalReferenceExists(currentPart,external);
                                if (!check || externalRef.Exists == false)
                                {
                                    if (externalRef.MatchPoint != null)
                                    {
                                        selectedCatiaProduct.Clear();
                                    
                                            selectedCatiaProduct.Add((AnyObject)product.Publications.Item(externalRef.MatchPoint).Valuation);
                                            selectedCatiaProduct.Copy();
                                            selectedCatiaProduct.Clear();
                                            selectedCatiaProduct.Add(currentPart);
                                            selectedCatiaProduct.PasteSpecial("CATPrtResult");
                                      
                                            selectedCatiaProduct.Clear();
                                       
                                        selectedCatiaProduct.Clear();
                                    }

                                }
                            }
                        }
                    }
                    break;
                }
 
            }
            return selectedCatiaProduct;
        }

        private static ExternalReferences CheckIfExternalReferenceExists(Part currentPart,ExternalReferences external)
        {
            foreach (HybridBody hybridBody in currentPart.HybridBodies)
            {
                if (hybridBody.get_Name() == "External References")
                {
                    check = true;
                    foreach (HybridShape hybridShape in hybridBody.HybridShapes)
                    {
                        if (external.MatchPoint == hybridShape.get_Name() || external.Point == hybridShape.get_Name())
                        {
                            external.Exists = true;
                            return external;
                        }
                    }
                }
            }
            return external;
        }
    }
    #region publication trial ideas
    /*  publications.Add(catiaProduct.);*/
    //get the skeleton and store it's publication as list                                                   
    //-------------------to note----------------                                                   
    //product                                                   
    //reference                                                  
    //CreateReferenceFromName                                                  
    //product.Publication                                                  
    //publication.Add("....some_name")                                                
    //publication.SetDirect("....some_name", reference)                                                  
    //foreach(Publication publication in product.Publications)                                                   
    //{                                                  
    //    publicationPoints.Add(publication.get_Name());                                                   
    //}
    #endregion

}