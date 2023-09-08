using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_GetCell : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_GetCell class.
        /// </summary>
        public GH_XL_Rng_GetCell()
          : base("Range Cell", "Rng Cell",
              "Returns specific cell",
              Constants.ShortName, Constants.SubRange)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Input, GH_ParamAccess.item);
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Status", "S", "The status of the selected cell. Returns true if the cell had contents and false if it was blank", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            DA.GetData(0, ref gooA);
            gooA.TryGetRange(out ExRange range);

            IGH_Goo gooB = null;
            DA.GetData(1, ref gooB);
            gooB.TryGetCell(out ExCell cell);

            bool status = range.TryGetCell(cell, out ExCell output);

            DA.SetData(0, output);
            DA.SetData(1, status);
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
                return Properties.Resources.XL_Rng_GetCell;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d0d2d31c-98f1-4f79-b096-994f796b1a2d"); }
        }
    }
}