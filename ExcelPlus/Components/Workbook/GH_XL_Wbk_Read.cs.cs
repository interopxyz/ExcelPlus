using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Workbook
{
    public class GH_XL_Wbk_Read : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wbk_Read class.
        /// </summary>
        public GH_XL_Wbk_Read()
          : base("Read Workbook", "Read Wbk",
              "Read a Workbook from a Memory Stream string",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.septenary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Stream", "S", "A string version of an excel file memory stream", GH_ParamAccess.item);
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Workbook.Name, Constants.Workbook.NickName, Constants.Workbook.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            bool activate = false;
            DA.GetData(1, ref activate);
            if (activate)
            {
                string stream = string.Empty;
                if (DA.GetData(0, ref stream))
                {
                    ExWorkbook workbook = new ExWorkbook();
                    workbook.Read(stream);
                    DA.SetData(0, workbook);
                }
            }

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
                return Properties.Resources.XL_Wbk_Read;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d7c666ef-d0a3-4921-af48-33eec41d2a9f"); }
        }
    }
}