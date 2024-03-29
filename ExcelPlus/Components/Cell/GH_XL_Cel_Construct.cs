﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public class GH_XL_Cel_Set : GH_XL_Cel__Base
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Cel_Set class.
        /// </summary>
        public GH_XL_Cel_Set()
          : base("Cell Content", "Cell",
              "Gets or Sets the contents of a Cell",
              Constants.ShortName, Constants.SubCell)
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
            pManager[0].Optional = true;
            pManager.AddGenericParameter(Constants.Location.Name, "Optional " + Constants.Location.NickName, Constants.Location.Input, GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager.AddTextParameter("Value", "V", "Optional Cell value", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddTextParameter(Constants.Format.Name, "Optional "+Constants.Format.NickName, Constants.Format.Input, GH_ParamAccess.item);
            pManager[3].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            base.RegisterOutputParams(pManager);
            pManager.AddPointParameter("Location", "L", "Cell location", GH_ParamAccess.item);
            ((IGH_PreviewObject)pManager[1]).Hidden = true;
            pManager.AddTextParameter("Value", "V", "Cell value", GH_ParamAccess.item);
            pManager.AddTextParameter("Format", "F", "Cell format", GH_ParamAccess.item);
            pManager.AddTextParameter("Formula", "=", "A cell formula if applicable", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooC = null;
            DA.GetData(0, ref gooC);
            gooC.TryGetCell(out ExCell cell);

            IGH_Goo gooL = null;
            if (DA.GetData(1, ref gooL)) if (gooL.TryGetCell(out ExCell location)) cell.Address = location.Address;

            string value = string.Empty;
            if (DA.GetData(2, ref value)) cell.Value = value;

            string format = string.Empty;
            if (DA.GetData(3, ref format)) cell.Format = format;

            DA.SetData(0, new ExCell(cell));
            DA.SetData(1, new Point3d(cell.Column,cell.Row,0));
            DA.SetData(2, cell.Value);
            DA.SetData(3, cell.Format);
            DA.SetData(4, cell.Formula);
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
                return Properties.Resources.XL_Cel_Construct;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("2f33480d-1374-4bb5-a8fb-5959b369204b"); }
        }
    }
}