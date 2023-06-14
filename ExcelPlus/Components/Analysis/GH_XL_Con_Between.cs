﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;

namespace ExcelPlus.Components.Analysis
{
    public class GH_XL_Con_Between : GH_XL_Con__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Con_Between class.
        /// </summary>
        public GH_XL_Con_Between()
          : base("Conditional Between", "Con Between",
              "Applies conditional formatting for values inside a bounds to a range",
              Constants.ShortName, Constants.SubAnalysis)
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
            pManager.AddIntervalParameter("Domain", "D", "The domain to evaluate", GH_ParamAccess.item);
            pManager.AddColourParameter("Cell Color", "C", "The cell highlight color", GH_ParamAccess.item, Sd.Color.LightGray);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Flip", "F", "If true, non unique values will be highlighted", GH_ParamAccess.item, false);
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
            IGH_Goo gooR = null;
            DA.GetData(0, ref gooR);
            gooR.TryGetRange(out ExRange range);

            Interval domain = new Interval(0,1);
            DA.GetData(1, ref domain);

            Sd.Color color1 = Constants.StartColor;
            DA.GetData(2, ref color1);

            bool flip = false;
            DA.GetData(3, ref flip);

            range.AddConditions(ExCondition.CreateBetweenCondition(domain.Min,domain.Max,flip,color1));
            
            DA.SetData(0, range);
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
                return Properties.Resources.XL_Con_Between;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("c7cc1156-2cc3-42a6-a631-0e7196a88656"); }
        }
    }
}