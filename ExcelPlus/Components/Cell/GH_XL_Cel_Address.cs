using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Cell
{
    public class GH_XL_Cel_Address : GH_XL_Cel__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Cel_Address class.
        /// </summary>
        public GH_XL_Cel_Address()
          : base("Cell Address", "Cell Addr",
              "Gets or Sets a Cell's Address",
              Constants.ShortName, Constants.SubCell)
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
            pManager[0].Optional = true;
            pManager.AddTextParameter("Address", "A", "The Cell Address", GH_ParamAccess.item);
            pManager[1].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Address", "A", "The Cell Address", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooC = null;
            DA.GetData(0, ref gooC);
            gooC.TryGetCell(out ExCell cell);

            string address = string.Empty;
            if (DA.GetData(1, ref address)) cell.Address = address;

            DA.SetData(0, cell);
            DA.SetData(1, cell.Address);
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
                return Properties.Resources.XL_Cel_Address;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("7462ed50-1b73-48a8-8fa9-41ad4b829e13"); }
        }
    }
}