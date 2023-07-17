using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_SubRanges : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_SubRanges class.
        /// </summary>
        public GH_XL_Rng_SubRanges()
          : base("Sub Ranges", "Sub Rng",
              "Returns a list of Ranges for each Column or Row",
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
            pManager.AddBooleanParameter("By Column", "C", "If true, Ranges will be returned by column. If false, by row", GH_ParamAccess.item, false);
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

            bool byColumn = false;
            DA.GetData(1, ref byColumn);

            DA.SetDataList(0, range.SubRanges(byColumn));
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
                return Properties.Resources.XL_Rng_By;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("3b0f3bfd-49bb-48a4-983a-4af403017335"); }
        }
    }
}