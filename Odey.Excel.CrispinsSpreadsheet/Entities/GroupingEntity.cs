﻿using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public abstract class GroupingEntity : IChildEntity
    {
        public GroupingEntity(string code, int firstRowOffset)
        {
            Code = code;
            _firstRowOffset = firstRowOffset;
        }

        private int _firstRowOffset;

        public string Code { get; private set; }

        public string Name { get; set; }

        public int FirstRowNumber => Children.Min(a => a.Value.RowNumber)- _firstRowOffset;

        public XL.Range TotalRow { get; set; }

        public int RowNumber => TotalRow.Row;

        public override string ToString()
        {
            return $"{Code}({Name})";
        }

        public Dictionary<string, IChildEntity> Children { get; set; } = new Dictionary<string, IChildEntity>();

        public string ControlString { get; set; }

        public decimal? Nav { get; set; }

        public decimal? PreviousNav { get; set; }
    }
}
