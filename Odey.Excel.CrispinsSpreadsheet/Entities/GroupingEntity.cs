using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public abstract class GroupingEntity : IChildEntity
    {
        public GroupingEntity(GroupingEntity parent,string code,string name,bool childrenArePositions,object ordering)
        {
            Identifier = new Identifier(null,code);
            Name = name;
            ChildrenArePositions = childrenArePositions;
            Ordering = ordering;
            Parent = parent;
        }

        public GroupingEntity Parent { get; private set; }

        public object Ordering { get; private set; }

        public Identifier Identifier { get; private set;}

        public string Name { get; set; }

        public GroupingEntity Previous { get; set; }

        public XL.Range TotalRow { get; set; }

        public int RowNumber => TotalRow.Row;

        public override string ToString()
        {
            return $"{Identifier.Code}({Name})";
        }

        public bool ChildrenArePositions { get; set; }

        public Dictionary<Identifier, IChildEntity> Children { get; set; } = new Dictionary<Identifier, IChildEntity>();

        public string ControlString { get; set; }

        public decimal? Nav { get; set; }

        public decimal? PreviousNav { get; set; }
    }
}
