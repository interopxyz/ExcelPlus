using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Workbook
{
    public class GH_XL_Wbk_Open : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wbk_Open class.
        /// </summary>
        public GH_XL_Wbk_Open()
          : base("Open Workbook", "Open Wbk",
              "Open a Workbook from an Excel file"+Environment.NewLine+ "Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.'",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.senary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Filepath", "P", "The full Filepath to an Excel file" + Environment.NewLine + "Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.'", GH_ParamAccess.item);
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
                string filepath = string.Empty;
                if (DA.GetData(0, ref filepath))
                {
                    ExWorkbook workbook = new ExWorkbook();
                    workbook.Open(filepath);
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
                return Properties.Resources.XL_Wbk_Open;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("0f4a8727-5d30-4a07-a4cc-279197f1a7b7"); }
        }
    }
}