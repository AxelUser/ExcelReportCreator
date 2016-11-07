using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ExcelReportCreator.Handler.Types
{
    public class CellStyle
    {
        public int CellsToMergeHorizontally { get; set; }
        public int CellsToMergeUpright { get; set; }
        public Color CellsColor { get; set; }
        public bool Border { get; set; }
        public bool BoldText { get; set; }
        public bool WordWrap { get; set; }

        public CellStyle()
        {
            CellsToMergeHorizontally = 1;
            CellsToMergeUpright = 1;
            WordWrap = true;
            BoldText = false;
        }
    }
}
