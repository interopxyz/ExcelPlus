﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace ExcelPlus.Components.Analysis
{
    public class GH_XL_Con_Top : GH_XL_Con__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Con_Top class.
        /// </summary>
        public GH_XL_Con_Top()
          : base("Conditional Top", "Con Top",
              "Applies top percent based conditional formatting to a Range",
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
            pManager.AddIntegerParameter("Total", "T", "The total number of values to highlight", GH_ParamAccess.item, 10);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Flip", "F", "If true, the bottom percent will be highlighted", GH_ParamAccess.item, false);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Cell Color", "C", "The cell highlight color", GH_ParamAccess.item, Constants.StartColor);
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

            int value = 0;
            DA.GetData(1, ref value);

            bool flip = false;
            DA.GetData(2, ref flip);

            Sd.Color color1 = Constants.StartColor;
            DA.GetData(3, ref color1);

            ExCondition condition = ExCondition.CreateTopCountCondition(value, flip, color1);

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
                return Properties.Resources.XL_Con_Count;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9dc0b177-0feb-4f72-937e-205e5085dfc4"); }
        }
    }
}