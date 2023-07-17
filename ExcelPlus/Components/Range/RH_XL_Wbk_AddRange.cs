using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class RH_XL_Wbk_AddRange : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the RH_XL_Wbk_SetRange class.
        /// </summary>
        public RH_XL_Wbk_AddRange()
          : base("New Range", "Rng New",
              "Creates a new Range from a minimum and maximum Cell",
              Constants.ShortName, Constants.SubRange)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Minimum Cell", "<", Constants.Cell.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter("Maximum Cell", ">", Constants.Cell.Input, GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            DA.GetData(0, ref gooA);
            gooA.TryGetCell(out ExCell min);

            IGH_Goo gooB = null;
            DA.GetData(1, ref gooB);
            gooB.TryGetCell(out ExCell max);

            DA.SetData(0, new ExRange(min,max));
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
                return Properties.Resources.XL_Rng_GetCells;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("debea2e6-7917-490c-ada8-dfe4d1843747"); }
        }
    }
}