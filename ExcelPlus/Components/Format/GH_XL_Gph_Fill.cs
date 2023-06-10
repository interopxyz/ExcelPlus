using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Gph_Fill : GH_XL_Gph__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Gph_Fill class.
        /// </summary>
        public GH_XL_Gph_Fill()
          : base("Fill", "Fill",
              "Set the Fill formatting for a cell or range",
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
            pManager.AddColourParameter("Object Color", "C", "Object color", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddColourParameter("Object Color", "C", "Object color", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            Color color = Color.Transparent;

            if(goo.CastTo<ExCell>(out ExCell cell))
            {
                cell = new ExCell(cell);
                if(DA.GetData(1,ref color)) cell.Graphic.FillColor = color;
                
                DA.SetData(0, cell);
                DA.SetData(1, cell.Graphic.FillColor);
            }
            else if(goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                if (DA.GetData(1, ref color)) range.Graphic.FillColor = color;

                DA.SetData(0, range);
                DA.SetData(1, range.Graphic.FillColor);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);
                if (DA.GetData(1, ref color)) sheet.Graphic.FillColor = color;

                DA.SetData(0, sheet);
                DA.SetData(1, sheet.Graphic.FillColor);
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
                return Properties.Resources.XL_Grp_Fill;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("ff8be311-c3c2-46d3-bb99-bd5d7cd5410f"); }
        }
    }
}