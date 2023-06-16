﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace ExcelPlus.Components
{
    public class GH_XL_Con_Value : GH_XL_Con__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Con_Value class.
        /// </summary>
        public GH_XL_Con_Value()
          : base("Conditional Value", "Con Value",
              "Applies value based conditional formatting to a range",
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
            pManager.AddNumberParameter("Value", "V", "The value to check against", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "The condition type", GH_ParamAccess.item, 0);
            pManager[2].Optional = true;
            pManager.AddColourParameter("Cell Color", "C", "The cell highlight color", GH_ParamAccess.item, Constants.StartColor);
            pManager[3].Optional = true;

            Param_Integer paramA = (Param_Integer)pManager[2];
            foreach (ValueCondition value in Enum.GetValues(typeof(ValueCondition)))
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
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            if (!DA.GetData(0, ref goo)) return;

            double value = 0.5;
            DA.GetData(1, ref value);

            int type = 0;
            DA.GetData(2, ref type);

            Sd.Color color1 = Constants.StartColor;
            DA.GetData(3, ref color1);

            ExCondition condition = ExCondition.CreateValueCondition(value, (ValueCondition)type, color1);

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
                return Properties.Resources.XL_Con_Value;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("be981ea3-33b4-4b78-8815-c2202c223f4f"); }
        }
    }
}