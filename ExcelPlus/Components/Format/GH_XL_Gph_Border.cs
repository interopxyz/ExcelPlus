using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Gph_Border : GH_XL_Gph__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Gph_BorderHorizontal class.
        /// </summary>
        public GH_XL_Gph_Border()
          : base("Border", "Border",
              "Get or Set the Border formatting for a Cell, Range, or Worksheet",
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
            pManager.AddIntegerParameter("Border", "B", "Border type", GH_ParamAccess.item,0);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Linetype", "L", "Border linetype", GH_ParamAccess.item);
            pManager[3].Optional = true;
            //pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
            //pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (Borders value in Enum.GetValues(typeof(Borders)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[3];
            foreach (LineTypes value in Enum.GetValues(typeof(LineTypes)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }

            //Param_Integer paramC = (Param_Integer)pManager[3];
            //foreach (BorderWeights value in Enum.GetValues(typeof(BorderWeights)))
            //{
            //    paramC.AddNamedValue(value.ToString(), (int)value);
            //}

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddIntegerParameter("Border", "B", "Border type", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Linetype", "L", "Border linetype", GH_ParamAccess.item);
            //pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            int borderType = 0;
            if (DA.GetData(1, ref borderType));
            Borders type = (Borders)borderType;

            Color color = Color.Black;
            int weight = 0;
            int linetype = 0;

            ExBorder border = new ExBorder();

            if (goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);

                switch(type)
                {
                    case Borders.Bottom:
                        border = cell.Graphic.BorderBottom;
                        break;
                    case Borders.Top:
                        border = cell.Graphic.BorderTop;
                        break;
                    case Borders.Left:
                        border = cell.Graphic.BorderLeft;
                        break;
                    case Borders.Right:
                        border = cell.Graphic.BorderRight;
                        break;
                    case Borders.Inside:
                        border = cell.Graphic.BorderInside;
                        break;
                    case Borders.Outside:
                        border = cell.Graphic.BorderOutside;
                        break;
                }

                if (DA.GetData(2, ref color)) border.Color = color;
                if (DA.GetData(3, ref linetype)) border.LineType = (LineTypes)linetype;

                DA.SetData(0, cell);
            }
            else if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);

                switch (type)
                {
                    case Borders.Bottom:
                        border = range.Graphic.BorderBottom;
                        break;
                    case Borders.Top:
                        border = range.Graphic.BorderTop;
                        break;
                    case Borders.Left:
                        border = range.Graphic.BorderLeft;
                        break;
                    case Borders.Right:
                        border = range.Graphic.BorderRight;
                        break;
                    case Borders.Inside:
                        border = range.Graphic.BorderInside;
                        break;
                    case Borders.Outside:
                        border = range.Graphic.BorderOutside;
                        break;
                }

                if (DA.GetData(2, ref color)) border.Color = color;
                if (DA.GetData(3, ref linetype)) border.LineType = (LineTypes)linetype;

                DA.SetData(0, range);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);

                switch (type)
                {
                    case Borders.Bottom:
                        border = sheet.Graphic.BorderBottom;
                        break;
                    case Borders.Top:
                        border = sheet.Graphic.BorderTop;
                        break;
                    case Borders.Left:
                        border = sheet.Graphic.BorderLeft;
                        break;
                    case Borders.Right:
                        border = sheet.Graphic.BorderRight;
                        break;
                    case Borders.Inside:
                        border = sheet.Graphic.BorderInside;
                        break;
                    case Borders.Outside:
                        border = sheet.Graphic.BorderOutside;
                        break;
                }

                if (DA.GetData(2, ref color)) border.Color = color;
                if (DA.GetData(3, ref linetype)) border.LineType = (LineTypes)linetype;

                DA.SetData(0, sheet);
            }
            else
            {
                return;
            }
                DA.SetData(1, type);
                DA.SetData(2, border.Color);
                DA.SetData(3, (int)border.LineType);
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
                return Properties.Resources.XL_Grp_BorderHorizontal;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("e4d7dc49-9750-43f9-ac30-3525d909ab31"); }
        }
    }
}