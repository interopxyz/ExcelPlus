using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Gph_Font : GH_XL_Gph__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Gph_Font class.
        /// </summary>
        public GH_XL_Gph_Font()
          : base("Font", "Font",
              "Set the Font formatting for a cell or range",
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
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager[4].Optional = true;
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager[5].Optional = true;
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
            pManager[6].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[4];
            foreach (Justifications value in Enum.GetValues(typeof(Justifications)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Family Name", "F", "Font Family name", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Font color", GH_ParamAccess.item);
            pManager.AddNumberParameter("Size", "S", "Font size", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Justification", "J", "Text justifications", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Bold", "B", "Font Bold status", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Italic", "I", "Font Italic status", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            string family = "Arial";
            Color color = Color.Black;
            double size = 10.0;
            int justifications = 1;
            bool isBold = false;
            bool isItalic = false;

            if (goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);
                if (DA.GetData(1, ref family)) cell.Font.Family = family;
                if (DA.GetData(2, ref color)) cell.Font.Color = color;
                if (DA.GetData(3, ref size)) cell.Font.Size = size;
                if (DA.GetData(4, ref justifications)) cell.Font.Justification = (Justifications)justifications;
                if (DA.GetData(5, ref isBold)) cell.Font.IsBold = isBold;
                if (DA.GetData(6, ref isItalic)) cell.Font.IsItalic = isItalic;

                DA.SetData(0, cell);
                DA.SetData(1, cell.Font.Family);
                DA.SetData(2, cell.Font.Color);
                DA.SetData(3, cell.Font.Size);
                DA.SetData(4, (int)cell.Font.Justification);
                DA.SetData(5, cell.Font.IsBold);
                DA.SetData(6, cell.Font.IsItalic);
            }
            else if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                if (DA.GetData(1, ref family)) range.Font.Family = family;
                if (DA.GetData(2, ref color)) range.Font.Color = color;
                if (DA.GetData(3, ref size)) range.Font.Size = size;
                if (DA.GetData(4, ref justifications)) range.Font.Justification = (Justifications)justifications;
                if (DA.GetData(5, ref isBold)) range.Font.IsBold = isBold;
                if (DA.GetData(6, ref isItalic)) range.Font.IsItalic = isItalic;

                DA.SetData(0, range);
                DA.SetData(1, range.Font.Family);
                DA.SetData(2, range.Font.Color);
                DA.SetData(3, range.Font.Size);
                DA.SetData(4, (int)range.Font.Justification);
                DA.SetData(5, range.Font.IsBold);
                DA.SetData(6, range.Font.IsItalic);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);
                if (DA.GetData(1, ref family)) sheet.Font.Family = family;
                if (DA.GetData(2, ref color)) sheet.Font.Color = color;
                if (DA.GetData(3, ref size)) sheet.Font.Size = size;
                if (DA.GetData(4, ref justifications)) sheet.Font.Justification = (Justifications)justifications;
                if (DA.GetData(5, ref isBold)) sheet.Font.IsBold = isBold;
                if (DA.GetData(6, ref isItalic)) sheet.Font.IsItalic = isItalic;

                DA.SetData(0, sheet);
                DA.SetData(1, sheet.Font.Family);
                DA.SetData(2, sheet.Font.Color);
                DA.SetData(3, sheet.Font.Size);
                DA.SetData(4, (int)sheet.Font.Justification);
                DA.SetData(5, sheet.Font.IsBold);
                DA.SetData(6, sheet.Font.IsItalic);
            }
            else
            {
                return;
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
                return Properties.Resources.XL_Grp_Font;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ca119a22-4646-47c8-bfdd-e88d548a9103"); }
        }
    }
}