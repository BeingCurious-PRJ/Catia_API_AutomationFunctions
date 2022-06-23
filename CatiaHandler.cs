using System;
using INFITF;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Create_CatiaMechanism;
using MECMOD;
using ProductStructureTypeLib;
using System.Runtime.InteropServices;

namespace Automation-Functions_Methods
{
    public static class CatiaHandler
    {
        #region Active Catia Check
        private static Application catia = null;
        public static Application Catia
        {
            get
            {
                if (catia == null)
                {
                    catia = GetCatiaObject();
                }
                return catia;
            }
            set => catia = value;
        }
        public static Application GetCatiaObject()
        {
            Application oCatia = null;
            try
            {
                oCatia = (Application)Marshal.GetActiveObject("CATIA.Application");
            }
            catch
            {
                throw new COMException("Error getting catia instance");
            }
            return oCatia;
        }
        #endregion 

        public static Part GetPartFromProduct(Product product)
        {
            //TODO use Try and make sure that we have the correct stuff --> this one here is NOT clean
            try
            {
                return (product.ReferenceProduct.Parent as PartDocument).Part;
            }
            catch
            {
                return null;
            }
        
        }

        public static Product GetProductByDefinition(Product product, string wfDefinition)
        {
            foreach (Product partProduct in product.Products)
            {
                if (partProduct.get_Definition() == wfDefinition)
                {
                    return partProduct;
                }
            }
            return null;
        }

        public static Product GetProductFromProductDocument(ProductDocument productDocument)
        {
            return productDocument.Product;
        }

        public static ProductDocument GetActiveProductDocument(Application catiaApp)
        {
            return (ProductDocument)catiaApp.ActiveDocument;
        }

        #region other usefull CatiaStuff
        private static Application _catiaAppObject;
        public static Application CatiaAppObject
        {
            set
            {
                _catiaAppObject = SetCatiaObject(value);
            }
            get
            {

                return GetCatiaObject(_catiaAppObject);
            }
        }
        public static Application SetCatiaObject(Application catiaApp = null)
        {
            Application resCatiaObj = null;
            if (catiaApp != null)
            {
                resCatiaObj = catiaApp;
            }
            return resCatiaObj;
        }
        public static Application GetCatiaObject(Application CatiaObj)
        {
            if (CatiaObj == null)
            {
                Type acType = Type.GetTypeFromProgID("CATIA.Application");
                CatiaObj = (Application)Activator.CreateInstance(acType, true);
            }
            return CatiaObj;

        }

        public static bool CloseAllDocuments(Application _catiaObj)
        {
            bool resVal = false;
            int docCount = _catiaObj.Documents.Count;

            for (int i = 1; i <= docCount; i++)
            {
                try
                {
                    _catiaObj.Documents.Item(i).Close();
                }
                catch (Exception)
                {

                    throw;
                }
            }
            //check if all docs are closed
            if (_catiaObj.Documents.Count == 0)
            {
                resVal = true;
            }
            return resVal;
        }
        public static Document[] GetAllDocuments(Application _catiaObj)
        {
            Documents allDocuments = _catiaObj.Documents;
            int docCount = allDocuments.Count;

            Document[] resAllDocuments = new Document[docCount];

            for (int i = 0; i < docCount; i++)
            {
                resAllDocuments[i] = allDocuments.Item(i);
            }

            return resAllDocuments;
        }
        #endregion
    }

}