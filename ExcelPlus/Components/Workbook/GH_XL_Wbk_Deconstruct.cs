﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components.Workbook
{
    public class GH_XL_Wbk_Deconstruct : GH_XL_Wbk__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Wbk_Deconstruct class.
        /// </summary>
        public GH_XL_Wbk_Deconstruct()
          : base("Deconstruct Workbook", "De Wbk",
              "Deconstruct a Workbook into its Worksheets",
              Constants.ShortName, Constants.SubWorkBooks)
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
            pManager.AddTextParameter("Names", "N", "Optional list of Worksheet names", GH_ParamAccess.list);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddGenericParameter(Constants.Worksheet.Name, Constants.Worksheet.NickName, Constants.Worksheet.Outputs, GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooB = null;
            DA.GetData(0, ref gooB);
            if (!gooB.TryGetWorkbook(out ExWorkbook workbook)) return;
            
            List<string> names = new List<string>();
            if(DA.GetDataList(1,names))
            {
                List<ExWorksheet> sheets = new List<ExWorksheet>();
                foreach (string name in names) if (workbook.TryGetSheet(name, out ExWorksheet sheet)) sheets.Add(sheet);
                DA.SetDataList(1, sheets);
            }
            else
            {
            DA.SetDataList(1, workbook.GetWorksheets());
            }
            

            DA.SetData(0, workbook);

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
                return Properties.Resources.XL_Wbk_Deconstruct;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("d43c1928-c84a-44c6-b694-fd562c6cad4e"); }
        }
    }
}