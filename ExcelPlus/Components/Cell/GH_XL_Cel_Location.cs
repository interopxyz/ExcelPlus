using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Cell
{
    public class GH_XL_Cel_Location : GH_XL_Cel__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Cel_Location class.
        /// </summary>
        public GH_XL_Cel_Location()
          : base("Cell Location", "Cell Loc",
              "Gets or Sets a Cell's Location",
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
            pManager.AddIntegerParameter("Column", "C", "The Cell column index", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddIntegerParameter("Row", "R", "The Cell row index", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Absolute Column", "AC", "Is the Cell column absolute", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddBooleanParameter("Absolute Row", "AR", "Is the Cell row absolute", GH_ParamAccess.item);
            pManager[4].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddIntegerParameter("Column", "C", "The Cell Column index", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Row", "R", "The Cell Row index", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Absolute Column", "AC", "Is the Cell Column absolute", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Absolute Row", "AR", "Is the Cell Row absolute", GH_ParamAccess.item);
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

            int column = 1;
            if (DA.GetData(1, ref column)) cell.Column = column;

            int row = 1;
            if (DA.GetData(2, ref row)) cell.Row = row;

            bool isColumn = false;
            if (DA.GetData(3, ref isColumn)) cell.IsColumnAbsolute = isColumn;

            bool isRow = false;
            if (DA.GetData(4, ref isRow)) cell.IsRowAbsolute = isRow;

            DA.SetData(0, cell);
            DA.SetData(1, cell.Column);
            DA.SetData(2, cell.Row);
            DA.SetData(3, cell.IsColumnAbsolute);
            DA.SetData(4, cell.IsRowAbsolute);
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
                return Properties.Resources.XL_Cel_Location;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("5b4657ff-78fb-4222-9535-8db699c5a20c"); }
        }
    }
}