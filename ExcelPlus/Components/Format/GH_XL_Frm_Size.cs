﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Frm_Size : GH_XL_Frm__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Frm_Size class.
        /// </summary>
        public GH_XL_Frm_Size()
          : base("Size", "Size",
              "Get or Set the Column and Row size for a Cell, Range, or Worksheet",
              Constants.ShortName, Constants.SubFormat)
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
            pManager.AddNumberParameter("Width", "W", "The Cell column width", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddNumberParameter("Height", "H", "The Cell row height", GH_ParamAccess.item);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddNumberParameter("Width", "W", "The Cell's column width", GH_ParamAccess.item);
            pManager.AddNumberParameter("Height", "H", "The Cell's row height", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            double width = -1;
            bool hasWidth = DA.GetData(1, ref width);

            double height = -1;
            bool hasHeight = DA.GetData(2, ref height);

            if (goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);
                if (hasWidth) cell.Width = width;
                if (hasHeight) cell.Height = height;
                DA.SetData(0, cell);
                DA.SetData(1, cell.Width);
                DA.SetData(2, cell.Height);
            }
            else if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                if (hasWidth) range.ColumnWidth = width;
                if (hasHeight) range.RowHeight= height;
                DA.SetData(0, range);
                DA.SetData(1, range.ColumnWidth);
                DA.SetData(2, range.RowHeight);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);
                if (hasWidth) sheet.ColumnWidth = width;
                if (hasHeight) sheet.RowHeight = height;
                DA.SetData(0, sheet);
                DA.SetData(1, sheet.ColumnWidth);
                DA.SetData(2, sheet.RowHeight);
            }
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
                return Properties.Resources.XL_Grp_Size;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("34337940-6d34-49ff-9382-137d20900b1f"); }
        }
    }
}