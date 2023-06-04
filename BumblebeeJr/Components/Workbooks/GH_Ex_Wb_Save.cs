using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using XL = ClosedXML;

namespace BumblebeeJr.Components.Workbooks
{
    public class GH_Ex_Wb_Save : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Wb_Save class.
        /// </summary>
        public GH_Ex_Wb_Save()
          : base("Save Workbook", "WbkSave",
              "Description",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Content", "T", "Temporary", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("FilePath", "P", "The filepath", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            XL.Excel.XLWorkbook workbook = new XL.Excel.XLWorkbook();
            XL.Excel.IXLWorksheet worksheet = workbook.AddWorksheet("Test");
            worksheet.Cell("A1").Value = "Hi";
            workbook.SaveAs("C:/Users/deman/Desktop/bbjr/test.xlsx");
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("92982d54-7809-477c-9fa4-677934a666cc"); }
        }
    }
}