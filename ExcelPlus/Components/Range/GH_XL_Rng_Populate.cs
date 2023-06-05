using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public class GH_XL_Rng_Populate : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Rng_Populate class.
        /// </summary>
        public GH_XL_Rng_Populate()
          : base("GH_XL_Rng_Populate", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Starting Cell", "C", Constants.Cell.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Input, GH_ParamAccess.item);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Output, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = null;
            DA.GetData(0, ref goo);
            if (!goo.TryGetCell(out ExCell cell)) return;

            GH_Structure<GH_String> ghData = new GH_Structure<GH_String>();

            List<List<GH_String>> dataSet = new List<List<GH_String>>();
            if (!DA.GetDataTree(2, out ghData)) return;

            foreach (List<GH_String> data in ghData.Branches)
            {
                dataSet.Add(data);
            }

            ExRange range = new ExRange(cell, dataSet);

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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("88560184-26a6-472c-884b-9d2b767f94e2"); }
        }
    }
}