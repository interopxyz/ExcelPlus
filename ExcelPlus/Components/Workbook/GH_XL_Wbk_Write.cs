using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Workbook
{
    public class GH_XL_Wbk_Write : GH_XL_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wbk_Write class.
        /// </summary>
        public GH_XL_Wbk_Write()
          : base("Write Workbook", "Write Workbook",
              "Write a Workbook to an Memory Stream",
              Constants.ShortName, Constants.SubWorkBooks)
        {
        }

        /// <summary>
        /// Set Exposure level for the component.
        /// </summary>
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.septenary; }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            base.RegisterInputParams(pManager);
            pManager.AddBooleanParameter(Constants.Activate.Name, Constants.Activate.NickName, Constants.Activate.Input, GH_ParamAccess.item);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddTextParameter("Stream", "S", "The memory stream", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            bool activate = false;
            DA.GetData(1, ref activate);
            if (activate)
            {
                IGH_Goo gooB = null;
                DA.GetData(0, ref gooB);
                if (!gooB.TryGetWorkbook(out ExWorkbook workbook)) return;

                string stream = workbook.Write();

                DA.SetData(0, workbook);
                DA.SetData(1, stream);
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
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("db01e4c2-daca-41a6-8b55-1024d386dbde"); }
        }
    }
}