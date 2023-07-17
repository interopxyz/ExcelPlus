using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public class GH_XL_Wbk_Properties : GH_XL_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wbk_Properties class.
        /// </summary>
        public GH_XL_Wbk_Properties()
          : base("Workbook Properties", "Wbk Prop",
              "Get or set Workbook properties",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddTextParameter("Name", "N", "The Workbook Name", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager); 
            pManager.AddTextParameter("Name", "N", "The Workbook Name", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooB = null;
            DA.GetData(0, ref gooB);
            if (!gooB.TryGetWorkbook(out ExWorkbook workbook)) return;

            string name = string.Empty;
            if (DA.GetData(1, ref name)) workbook.Name = name;

            DA.SetData(0, workbook);
            DA.SetData(1, workbook.Name);
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
                return Properties.Resources.XL_Wbk_Edit;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("30b99ea0-9fa7-49bc-8594-1f85d6d58051"); }
        }
    }
}