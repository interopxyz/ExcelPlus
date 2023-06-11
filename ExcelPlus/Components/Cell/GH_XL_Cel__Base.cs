﻿using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace ExcelPlus.Components
{
    public abstract class GH_XL_Cel__Base : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_XL_Cel__Base class.
        /// </summary>
        public GH_XL_Cel__Base()
          : base("GH_XL_Cel__Base", "Nickname",
              "Description",
              "Category", "Subcategory")
        {
        }

        public GH_XL_Cel__Base(string Name, string NickName, string Description, string Category, string Subcategory) : base(Name, NickName, Description, Category, Subcategory)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter(Constants.Cell.Name, Constants.Cell.NickName, Constants.Cell.Input, GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
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
            get { return new Guid("918d2799-f2b7-4246-8bc7-aa02ee1d40c3"); }
        }
    }
}