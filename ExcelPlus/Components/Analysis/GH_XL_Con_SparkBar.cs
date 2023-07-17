using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

using Sd = System.Drawing;
namespace ExcelPlus.Components.Analysis
{
    public class GH_XL_Con_SparkBar : GH_XL_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Con_SparkBar class.
        /// </summary>
        public GH_XL_Con_SparkBar()
          : base("Add Spark Bar", "SparkBar",
              "Adds a SparkBar to a Range",
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
            pManager.AddGenericParameter("Source Range", "Rng", Constants.Range.Input, GH_ParamAccess.item);
            pManager.AddGenericParameter("Location Range", "L", Constants.Range.Input, GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "The Spark Bar color", GH_ParamAccess.item, Constants.StartColor);
            pManager[3].Optional = true;
            pManager.AddNumberParameter("Weight", "W", "The Spark Bar weight", GH_ParamAccess.item, 1.0);
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
            IGH_Goo gooS = null;
            DA.GetData(0, ref gooS);
            if (!gooS.TryGetWorksheet(out ExWorksheet worksheet)) return;

            IGH_Goo gooR = null;
            DA.GetData(1, ref gooR);
            if (!gooR.TryGetRange(out ExRange range)) return;

            IGH_Goo gooL = null;
            DA.GetData(2, ref gooL);
            if (!gooL.TryGetRange(out ExRange location)) return;

            Sd.Color color = Constants.StartColor;
            DA.GetData(3, ref color);

            double weight = 1.0;
            DA.GetData(4, ref weight);

            worksheet.AddSparkLines(ExSpark.ConstructBar(location, range, color, weight));

            DA.SetData(0, worksheet);
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
                return Properties.Resources.XL_Con_SparkBar;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("3a8eb5e0-4c99-4080-b1db-5434a670f0bb"); }
        }
    }
}