using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Cell
{
    public class GH_XL_Cel_Cells : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Cel_Cells class.
        /// </summary>
        public GH_XL_Cel_Cells()
          : base("Generate Cells", "Cells",
              "Generates a list of Cells from a minimum Cell and maximum Cell",
              Constants.ShortName, Constants.SubCell)
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
            pManager.AddGenericParameter(Constants.Cell.Name, "<", Constants.Cell.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Cell.Name, ">", Constants.Cell.Input, GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Flip", "F", "If true, cells are listed by column. If false, by row", GH_ParamAccess.item, true);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, "C", Constants.Cell.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            DA.GetData(0, ref gooA);
            gooA.TryGetCell(out ExCell cellA);

            IGH_Goo gooB = null;
            DA.GetData(1, ref gooB);
            gooB.TryGetCell(out ExCell cellB);

            ExRange range = new ExRange(new List<ExCell> { cellA, cellB });

            bool flip = false;
            DA.GetData(2, ref flip);

            DA.SetDataList(0, range.Cells(flip));
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
                return Properties.Resources.XL_Cel_GetCells2;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("50fbdf6c-ca47-41eb-a577-a4de236d9c58"); }
        }
    }
}