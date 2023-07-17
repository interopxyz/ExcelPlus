using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_Merge : GH_XL_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_Merge class.
        /// </summary>
        public GH_XL_Rng_Merge()
          : base("Merge Range", "Merge Rng",
              "Merge or unmerge a Range"+Environment.NewLine+ "(note: Ranges will be merged sequentially. If multiple Ranges overlap, the latter Range merge will be ignored.)",
              Constants.ShortName, Constants.SubRange)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quarternary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddBooleanParameter("Merge", "M", "If true, the Cells in the Range will be merged and only the first Cell's value and formatting will be applied.", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddBooleanParameter("Merged", "M", "The merge status of the Range", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooR = null;
            DA.GetData(0, ref gooR);
            gooR.TryGetRange(out ExRange range);

            bool merged = false;
            if(DA.GetData(1, ref merged))range.IsMerged = merged;

            DA.SetData(0, range);
            DA.SetData(1, range.IsMerged);
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
                return Properties.Resources.XL_Rng_Merge;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("a8cfda30-23bd-479c-84e4-7e9dbeff2de9"); }
        }
    }
}