using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Range
{
    public class GH_XL_Rng_Deconstruct : GH_XL_Rng__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_Cells class.
        /// </summary>
        public GH_XL_Rng_Deconstruct()
          : base("Deconstruct Range", "De Rng",
              "Deconstruct a Range into it's Cells",
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
            base.RegisterInputParams(pManager);
            pManager.AddBooleanParameter("Active", "A", "Return only active cells. If true Flip (F) input will be ignored", GH_ParamAccess.item, false);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Flip", "F", "If true, cells are listed by column. If false, by row", GH_ParamAccess.item, true);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Outputs, GH_ParamAccess.list);
            pManager.AddTextParameter("Address", "A", "The Range address", GH_ParamAccess.item);
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

            bool flip = false;
            DA.GetData(2, ref flip);

            DA.SetData(0, range);

            if (active)
            {
                DA.SetDataList(1, range.ActiveCells);
            }
            else
            {
                DA.SetDataList(1, range.Cells(flip));
            }

            DA.SetData(2, range.Address);
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
                return Properties.Resources.XL_Rng_Deconstruct;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1e8b3c0e-dc75-4743-a50c-8fa0d6a6c8c0"); }
        }
    }
}