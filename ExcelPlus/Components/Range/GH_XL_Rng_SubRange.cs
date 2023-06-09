using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_SubRange :GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_SubRange class.
        /// </summary>
        public GH_XL_Rng_SubRange()
          : base("Sub Range", "Sub Rng",
              "Converts all Cells in a Range into a Range with a single Cell",
              Constants.ShortName, Constants.SubRange)
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
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Input, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Active", "A", "If true, only the active cells in the range will be returned", GH_ParamAccess.item, false);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Outputs, GH_ParamAccess.list);
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

            bool active = false;
            DA.GetData(1, ref active);

            DA.SetData(0, range);
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
            get { return new Guid("e9699125-bccf-4450-a604-ebf85f2d6b31"); }
        }
    }
}