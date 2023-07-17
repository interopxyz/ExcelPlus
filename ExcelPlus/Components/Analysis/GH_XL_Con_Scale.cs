using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace ExcelPlus.Components.Analysis
{
    public class GH_XL_Con_Scale : GH_XL_Con__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Con_Scale class.
        /// </summary>
        public GH_XL_Con_Scale()
          : base("Conditional Scale", "Con Scl",
              "Applies value scale conditional formatting to a range",
              Constants.ShortName, Constants.SubAnalysis)
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
            pManager.AddNumberParameter("Parameter", "P", "The parameter of the midpoint of a 3 color gradient", GH_ParamAccess.item, 0.5);
            pManager[1].Optional = true;
            pManager.AddColourParameter("Gradient Color 1", "C0", "The first color of the gradient", GH_ParamAccess.item, Constants.StartColor);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Gradient Color 2", "C1", "The second color of the gradient", GH_ParamAccess.item, Constants.MidColor);
            pManager[3].Optional = true;
            pManager.AddColourParameter("Gradient Color 3", "C2", "The third color of the gradient", GH_ParamAccess.item);
            pManager[4].Optional = true;
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

            double param = 0.5;
            DA.GetData(1, ref param);

            Sd.Color color1 = Constants.StartColor;
            DA.GetData(2, ref color1);

            Sd.Color color2 = Constants.MidColor;
            DA.GetData(3, ref color2);

            Sd.Color color3 = Constants.EndColor;
            bool hasMid = DA.GetData(4, ref color3);

            ExCondition condition = new ExCondition();

            if (hasMid)
            {
                condition = ExCondition.CreateScalarCondition(color1,color2,param,color3);
            }
            else
            {
                condition = ExCondition.CreateScalarCondition(color1,color2);
            }

            if (goo.CastTo<ExRange>(out ExRange range))
            {
                range = new ExRange(range);
                range.AddConditions(condition);
                DA.SetData(0, range);
            }
            else if (goo.CastTo<ExWorksheet>(out ExWorksheet sheet))
            {
                sheet = new ExWorksheet(sheet);
                sheet.AddConditions(condition);
                DA.SetData(0, sheet);
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
                return Properties.Resources.XL_Con_Scale;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d6d6d342-1607-4f56-9282-4cd70b8892b7"); }
        }
    }
}