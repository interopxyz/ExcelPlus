using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_Construct : GH_XL_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_Compile class.
        /// </summary>
        public GH_XL_Rng_Construct()
          : base("Construct Range", "Range",
              "Creates a Range from a series of Cells",
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
            base.RegisterInputParams(pManager);
            pManager[0].Optional = true;
            pManager.AddGenericParameter("Cells", "Cel", Constants.Cell.Input, GH_ParamAccess.list);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
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

            List<IGH_Goo> goos = new List<IGH_Goo>();
            DA.GetDataList(1, goos);
            if (goos.TryCompileRange(out ExRange rng)) range.SetCells(rng.ActiveCells);

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
                return Properties.Resources.XL_Rng_Add;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("8719fc09-3dca-4b8c-a3cd-34f6970f93c8"); }
        }
    }
}