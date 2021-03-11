using System;
using Inventor;

namespace InventorEdge
{
    internal class InventorController
    {
        public static Inventor.Application iApp = null;

        internal static void generaPE()
        {
            getIstance();
            
            importaFromAutocad("W0200669");
            
            generaProfilo("testtttttt", 385);

            settoPiani(4.5, 4.5, 385);

            inseriscoOP("OP1001");
        }
        
        private static void inseriscoOP(string v)
        {
            throw new NotImplementedException();
        }

        private static void settoPiani(double v1, double v2, double modulo)
        {
            PartDocument ThisDoc = (PartDocument)iApp.ActiveDocument;

            PartComponentDefinition oPartCompDef = ThisDoc.ComponentDefinition;

            SurfaceBody osb = oPartCompDef.SurfaceBodies[1];

            WorkPlane mMinore = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], 0);
            mMinore.Name = "AsseMinore";

            WorkPlane mMaggiore = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], modulo);
            mMaggiore.Name = "AsseMaggiore";

            WorkPlane taglioTop = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], modulo-v2);
            taglioTop.Name = "taglioTop";

            WorkPlane taglioBottom = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], v1);
            taglioBottom.Name = "taglioBottom";
        }

        private static void generaProfilo(string v, double modulo)
        {
            PartDocument ThisDoc = (PartDocument)iApp.ActiveDocument;

            PartComponentDefinition oPartCompDef = ThisDoc.ComponentDefinition;

            PlanarSketch oSketch = (PlanarSketch)oPartCompDef.Sketches[v];

            Profile oProfile = oSketch.Profiles.AddForSolid();

            ExtrudeFeature oExtrude = oPartCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, modulo, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
        }

        private static void importaFromAutocad(string v)
        {
            TranslatorAddIn oDWGTranslator = (TranslatorAddIn)iApp.ApplicationAddIns.ItemById["{C24E3AC2-122E-11D5-8E91-0010B541CD80}"];

            DataMedium oDataMedium = iApp.TransientObjects.CreateDataMedium();
            
            oDataMedium.FileName = @"X:\Commesse\Focchi\200000 40L\99 Service\Matrici full\"+v+"_def.dwg";

            TranslationContext oTranslationContext = iApp.TransientObjects.CreateTranslationContext();

            oTranslationContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;


            PartDocument oDoc = (PartDocument)iApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject);

            PartComponentDefinition oPartCompDef = oDoc.ComponentDefinition;

            PlanarSketch oSketchC = oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[2], true);

            oSketchC.Name = "testtttttt";

            Sketch oSketch = (Sketch)oPartCompDef.Sketches["testtttttt"];
            oSketch.Edit();

            oTranslationContext.OpenIntoExisting = oSketch;

            NameValueMap oOptions = iApp.TransientObjects.CreateNameValueMap();

            oOptions.Add("SelectedLayers", "0");

            oOptions.Add("InvertLayersSelection", false);
            oOptions.Add("ConstrainEndPoints", true);

            object boh;

            oDWGTranslator.Open(oDataMedium, oTranslationContext, oOptions, out boh);
            oSketch.ExitEdit();
            // RegisterAutoCADDefault()
        }

        private static void getIstance()
        {
            try
            {
                iApp = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                iApp.Visible = true;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                iApp = (Inventor.Application)System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Inventor.Application"));
                iApp.Visible = true;
            }
        }
    }
}