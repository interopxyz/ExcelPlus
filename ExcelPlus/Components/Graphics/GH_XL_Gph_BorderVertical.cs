using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components.Graphics
{
    public class GH_XL_Gph_BorderVertical : GH_XL_Gph__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Gph_Border class.
        /// </summary>
        public GH_XL_Gph_BorderVertical()
          : base("Vertical Border", "V Border",
              "Set the Vertical Border formatting for a cell or range",
              Constants.ShortName, Constants.SubGraphics)
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
            pManager.AddIntegerParameter("Border", "B", "Border type", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
            pManager[3].Optional = true;
            pManager.AddIntegerParameter("Linetype", "L", "Border linetype", GH_ParamAccess.item);
            pManager[4].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[1];
            foreach (VerticalBorder value in Enum.GetValues(typeof(VerticalBorder)))
            {
                paramA.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramB = (Param_Integer)pManager[3];
            foreach (BorderWeight value in Enum.GetValues(typeof(BorderWeight)))
            {
                paramB.AddNamedValue(value.ToString(), (int)value);
            }

            Param_Integer paramC = (Param_Integer)pManager[4];
            foreach (LineType value in Enum.GetValues(typeof(LineType)))
            {
                paramC.AddNamedValue(value.ToString(), (int)value);
            }

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddIntegerParameter("Border", "B", "Border type", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Weight", "W", "Border weight", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Linetype", "L", "Border linetype", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            int border = 0;
            Color color = Color.Black;
            int weight = 0;
            int linetype = 0;

            if (goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);
                if (DA.GetData(1, ref border)) cell.Graphic.VerticalBorder = (VerticalBorder)border;
                if (DA.GetData(2, ref color)) cell.Graphic.VerticalStrokeColor = color;
                if (DA.GetData(3, ref weight)) cell.Graphic.VerticalWeight = (BorderWeight)weight;
                if (DA.GetData(4, ref linetype)) cell.Graphic.VerticalType = (LineType)linetype;

                DA.SetData(0, cell);
                DA.SetData(1, (int)cell.Graphic.VerticalBorder);
                DA.SetData(2, cell.Graphic.VerticalStrokeColor);
                DA.SetData(3, (int)cell.Graphic.VerticalWeight);
                DA.SetData(4, (int)cell.Graphic.VerticalType);
            }
            else if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                if (DA.GetData(1, ref border)) range.Graphic.VerticalBorder = (VerticalBorder)border;
                if (DA.GetData(2, ref color)) range.Graphic.VerticalStrokeColor = color;
                if (DA.GetData(3, ref weight)) range.Graphic.VerticalWeight = (BorderWeight)weight;
                if (DA.GetData(4, ref linetype)) range.Graphic.VerticalType = (LineType)linetype;

                DA.SetData(0, cell);
                DA.SetData(1, (int)range.Graphic.VerticalBorder);
                DA.SetData(2, range.Graphic.VerticalStrokeColor);
                DA.SetData(3, (int)range.Graphic.VerticalWeight);
                DA.SetData(4, (int)range.Graphic.VerticalType);
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
                return Properties.Resources.XL_Grp_BorderVertial;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("761efeeb-a64b-460e-b104-35d073c0998c"); }
        }
    }
}