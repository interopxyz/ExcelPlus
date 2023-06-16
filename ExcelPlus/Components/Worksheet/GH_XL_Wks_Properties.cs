using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public class GH_XL_Wks_Properties : GH_XL_Wks__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wks_Properties class.
        /// </summary>
        public GH_XL_Wks_Properties()
          : base("Worksheet Properties", "Wks Prop",
              "Gets or Sets Worksheet properties",
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
            pManager.AddTextParameter("Name", "N", "The Worksheet Name", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddBooleanParameter("Active", "A", "Flag the Worksheet as active" + Environment.NewLine + "(note: If this is true on multiple worksheets, only the last sheet in the sequence will be set to active.)", GH_ParamAccess.item);
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Name", "N", "The Worksheet Name", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Active", "A", "Gets the active status of a Worksheet", GH_ParamAccess.item);
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

            string name = string.Empty;
            if (DA.GetData(1, ref name)) worksheet.Name = name;

            bool active = false;
            if (DA.GetData(2, ref active)) worksheet.Active = active;

            DA.SetData(0, worksheet);
            DA.SetData(1, worksheet.Name);
            DA.SetData(2, worksheet.Active);
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
                return Properties.Resources.XL_Wks_Edit;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("77c44219-2cc1-454f-935b-06a0487b6484"); }
        }
    }
}