using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = ClosedXML.Excel;

namespace ExcelPlus
{
    public class ExWorksheet
    {

        #region members

        public XL.IXLWorksheet ComObj = null;

        protected string name = string.Empty;

        protected ExRange baseRange = new ExRange();

        public List<ExRange> Ranges = new List<ExRange>();

        protected bool active = false;

        public ExGraphic Graphic = new ExGraphic();
        public ExFont Font = new ExFont();

        protected List<ExSpark> sparkLines = new List<ExSpark>();

        #endregion

        #region constructors

        public ExWorksheet()
        {
        }

        public ExWorksheet(XL.IXLWorksheet sheet)
        {
            this.name = sheet.Name;
            this.Ranges.Add(new ExRange(sheet.CellsUsed().ToList()));
            this.baseRange = new ExRange(sheet.Cells().ToList());

            this.baseRange.ColumnWidth = sheet.ColumnWidth;
            this.baseRange.RowHeight = sheet.RowHeight;

            this.Graphic = new ExGraphic(sheet.Style);
            this.Font = new ExFont(sheet.Style);

            this.active = sheet.TabActive;
        }

        public ExWorksheet(string name)
        {
            this.name = name;
        }

        public ExWorksheet(ExRange range)
        {
            this.Ranges.Add(new ExRange(range));
        }

        public ExWorksheet(List<ExRange> ranges)
        {
            foreach (ExRange range in ranges)
            {
                this.Ranges.Add(new ExRange(range));
            }
        }

        public ExWorksheet(List<ExCell> cells)
        {
            this.Ranges.Add(new ExRange(cells));
        }

        public ExWorksheet(ExWorksheet worksheet)
        {
            this.name = worksheet.name;
            this.active = worksheet.active;
            foreach(ExRange range in worksheet.Ranges)
            {
                this.Ranges.Add(new ExRange(range));
            }
            this.baseRange = worksheet.baseRange;

            this.Graphic = new ExGraphic(worksheet.Graphic);
            this.Font = new ExFont(worksheet.Font);

            this.baseRange.Conditions = worksheet.baseRange.Conditions;
            this.sparkLines = worksheet.SparkLines;
        }

        #endregion

        #region properties

        public List<ExSpark> SparkLines
        {
            get
            {
                List<ExSpark> output = new List<ExSpark>();
                foreach (ExSpark spark in sparkLines)
                {
                    output.Add(new ExSpark(spark));
                }
                return output;
            }
            set
            {
                List<ExSpark> output = new List<ExSpark>();
                foreach (ExSpark spark in value)
                {
                    output.Add(new ExSpark(spark));
                }
                sparkLines = output;
            }
        }

        public virtual string Name
        {
            get { return name; }
            set
            {
                this.name = value;
                if (this.ComObj != null) this.ComObj.Name = this.name;
            }
        }

        public virtual bool Active
        {
            get { return active; }
            set { active = value; }
        }

        public virtual List<ExCell> ActiveCells
        {
            get
            {
                List<ExCell> cells = new List<ExCell>();
                foreach (ExRange range in this.Ranges) cells.AddRange(range.ActiveCells);
                return cells;
            }
        }

        public virtual double ColumnWidth
        {
            get { return this.baseRange.ColumnWidth; }
            set { this.baseRange.ColumnWidth = value; }
        }

        public virtual double RowHeight
        {
            get { return this.baseRange.RowHeight; }
            set { this.baseRange.RowHeight = value; }
        }

        public virtual List<ExCondition> Conditions
        {
            get { return this.baseRange.Conditions; }
            set { this.baseRange.Conditions = value; }
        }

        #endregion

        #region methods

        public void AddSparkLines(ExSpark sparkLine)
        {
            this.sparkLines.Add(new ExSpark(sparkLine));
        }

        public void AddSparkLines(List<ExSpark> sparkLines)
        {
            foreach (ExSpark sparkLine in sparkLines) this.sparkLines.Add(new ExSpark(sparkLine));
        }

        public void ClearSparklines()
        {
            this.sparkLines = new List<ExSpark>();
        }

        public void AddConditions(ExCondition condition)
        {
            this.baseRange.AddConditions(condition);
        }

        public void AddConditions(List<ExCondition> conditions)
        {
            this.baseRange.AddConditions(conditions);
        }

        public void ClearConditions()
        {
            this.baseRange.ClearConditions();
        }

        public void ClearValues()
        {
            foreach(ExRange range in Ranges) range.ClearValues();
        }

        public void ClearFormatting()
        {
            this.Graphic = new ExGraphic();
            this.Font = new ExFont();

            this.ColumnWidth = -1;
            this.RowHeight = -1;

            foreach (ExRange range in Ranges) range.ClearFormatting();
        }

        public List<ExRange> GetRanges(bool collapse = false)
        {
            List<ExRange> ranges = new List<ExRange>();
            if (collapse)
            {
                ranges.Add(new ExRange(this.ActiveCells));
            }
            else
            {
                foreach (ExRange range in this.Ranges) ranges.Add(new ExRange(range));
            }

            return ranges;
        }

        public bool TryGetRange(int index, out ExRange result)
        {
            if (index>this.Ranges.Count)
            {
                result = new ExRange();
                return true;
            }
            else
            {
                result = new ExRange(this.Ranges[index]);
                return true;
            }
        }

        #endregion

        #region application

        public void ApplyFormatting(XL.IXLWorksheet input)
        {
            this.Graphic.Apply(input.Style);
            this.Font.Apply(input.Style);

            int c = this.baseRange.Conditions.Count;
            for (int i = 0; i < c; i++)
            {
                this.baseRange.Conditions[c - 1 - i].ApplyCondition(input.AddConditionalFormat());
            }

            foreach(ExSpark spark in this.sparkLines)
            {
                XL.IXLSparklineGroup xlSpark = input.SparklineGroups.Add(spark.Location.Address, spark.Range.Address);
                xlSpark.Style.SeriesColor = spark.Color.ToExcel();
                xlSpark.SetLineWeight(spark.Weight);
                if(spark.IsColumn)xlSpark.Type = XL.XLSparklineType.Column;
            }
        }


        #endregion

        #region overrides

        public override string ToString()
        {
            return "Worksheet | " + this.Name;
        }

        #endregion

    }
}
