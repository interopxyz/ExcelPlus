using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public class GH_XL_Wks_Deconstruct : GH_XL_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wks_GetCells class.
        /// </summary>
        public GH_XL_Wks_Deconstruct()
          : base("Deconstruct Worksheet", "De Wks",
              "Deconstruct a Worksheet into it's Ranges and active Cells",
              Constants.ShortName, Constants.SubWorkSheets)
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
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Range.Name, Constants.Range.NickName, Constants.Range.Outputs, GH_ParamAccess.list);
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Outputs, GH_ParamAccess.list);
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

            DA.SetData(0, worksheet);
            DA.SetDataList(1, worksheet.GetRanges());
            DA.SetDataList(2, worksheet.ActiveCells);
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
                return Properties.Resources.XL_Wks_Deconstruct;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("1dec7015-73fc-4501-9454-10f03c08e70c"); }
        }
    }
}