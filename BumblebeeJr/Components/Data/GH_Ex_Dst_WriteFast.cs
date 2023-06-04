using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using XL = ClosedXML.Excel;

namespace BumblebeeJr.Components.Data
{
    public class GH_Ex_Dst_WriteFast : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Dst_WriteFast class.
        /// </summary>
        public GH_Ex_Dst_WriteFast()
          : base("GH_Ex_Dst_WriteFast", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
            pManager[0].Optional = true;
            pManager.AddTextParameter("Values", "V", "A datatree of values", GH_ParamAccess.tree);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            XL.XLWorkbook workbook = new XL.XLWorkbook();
            XL.IXLWorksheet worksheet = workbook.AddWorksheet("Test");


            IGH_Goo gooC = null;
            DA.GetData(0, ref gooC);
            XLCell cell = new XLCell();
            gooC.TryGetCell(ref cell);

            GH_Structure<GH_String> ghData = new GH_Structure<GH_String>();
            List<List<GH_String>> dataSet = new List<List<GH_String>>();
            if (!DA.GetDataTree(1, out ghData)) return;

            foreach (List<GH_String> data in ghData.Branches)
            {
                dataSet.Add(data);
            }

            int y = dataSet[0].Count;
            int x = dataSet.Count;

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    worksheet.Cell(cell.Row+j,cell.Column+i).Value = dataSet[i][j].Value;
                }
            }

            workbook.SaveAs("C:/Users/deman/Desktop/bbjr/testdata.xlsx");
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
            get { return new Guid("1e92e41f-1d46-4088-beed-afd83ab50501"); }
        }
    }
}