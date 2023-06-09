﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Frm_Clear : GH_XL_Frm__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Frm_Clear class.
        /// </summary>
        public GH_XL_Frm_Clear()
          : base("Clear", "Clear",
              "Clear the content or formatting for a Cell or all Cells in a Range or Worksheet",
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
            pManager.AddBooleanParameter("Clear Values", "V", "Clear all values", GH_ParamAccess.item, false);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Clear Formatting", "F", "Clear all graphic formatting", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Clear Conditions", "C", "Clear all conditional formatting", GH_ParamAccess.item, false);
            pManager[3].Optional = true;
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
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            bool clearValue = false;
            DA.GetData(1, ref clearValue);

            bool clearFormat = false;
            DA.GetData(2, ref clearFormat);

            if (goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);
                if (clearValue) cell.ClearValue();
                if (clearFormat) cell.ClearFormatting();

                DA.SetData(0, cell);
            }
            else if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                if (clearValue) range.ClearValues();
                if (clearFormat) range.ClearFormatting();

                DA.SetData(0, range);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);
                if (clearValue) sheet.ClearValues();
                if (clearFormat) sheet.ClearFormatting();

                DA.SetData(0, sheet);
            }
            else if (goo.CastTo<ExWorkbook>(out ExWorkbook book))
            {
                book = new ExWorkbook(book);
                if (clearValue) book.ClearValues();
                if (clearFormat) book.ClearFormatting();

                DA.SetData(0, book);
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
                return Properties.Resources.XL_Cel_Clear;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("41370c2c-ebab-413e-bb81-dc3f14a033a9"); }
        }
    }
}