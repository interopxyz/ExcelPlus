using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExWorkbook
    {

        #region members

        public XL.IXLWorkbook ComObj = null;

        protected string name = "";

        public List<ExWorksheet> Sheets = new List<ExWorksheet>();

        #endregion

        #region constructors

        public ExWorkbook()
        {
            this.ComObj = new XL.XLWorkbook();
        }

        public ExWorkbook(string name)
        {
            this.name = name;
        }

        public ExWorkbook(XL.IXLWorkbook comObj)
        {
            this.ComObj = comObj;
            this.name = comObj.Properties.Title;
        }

        public ExWorkbook(ExWorkbook workbook)
        {
            this.ComObj = workbook.ComObj;
            this.name = workbook.Name;
            foreach (ExWorksheet sheet in workbook.Sheets)
            {
                this.Sheets.Add(new ExWorksheet(sheet));
            }
        }

        public ExWorkbook(List<ExWorksheet> sheets)
        {
            foreach (ExWorksheet sheet in sheets)
            {
                this.Sheets.Add(new ExWorksheet(sheet));
            }
        }

        public ExWorkbook(ExWorksheet sheet)
        {
                this.Sheets.Add(new ExWorksheet(sheet));
        }

        public ExWorkbook(List<ExRange> ranges)
        {
                this.Sheets.Add(new ExWorksheet(ranges));
        }

        public ExWorkbook(List<ExCell> cells)
        {
            this.Sheets.Add(new ExWorksheet(new List<ExRange> { new ExRange(cells)}));
        }


        #endregion

        #region properties

        public virtual string Name
        {
            get { return this.name; }
            set 
            { 
                ComObj.Properties.Title = value;
                this.name = value;
            }
        }

        #endregion

        #region methods

        public bool TryGetSheet(int index, out ExWorksheet result)
        {
            if (index > this.Sheets.Count)
            {
                result = new ExWorksheet();
                return true;
            }
            else
            {
                result = new ExWorksheet(this.Sheets[index]);
                return true;
            }
        }

        public string Save(string directory, Extensions extension)
        {
            if(name == string.Empty)
            {
                return this.Save(directory, Constants.UniqueName, extension);
            }
            else
            {
                return this.Save(directory, name, extension);
            }
        }

        public string Save(string directory, string filename, Extensions extension )
        {
            string folder = Path.GetDirectoryName(directory);
            string name = Path.GetFileNameWithoutExtension(filename);
            string filepath = folder + "/" + name + "."+extension.ToString();

            this.name = name;

            foreach(ExWorksheet sheet in Sheets)
            {
                XL.IXLWorksheet xlSheet = this.ComObj.AddWorksheet();
                if (sheet.Name != string.Empty) xlSheet.Name = sheet.Name;
                List<ExCell> cells = sheet.ActiveCells;
                foreach(ExCell cell in cells)xlSheet.SetValue()
            }

            this.ComObj.SaveAs(filepath);

            return filepath;
        }

        #endregion

    }
}
