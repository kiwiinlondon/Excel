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
        public GroupingEntity(GroupingEntity parent,string code,string name, EntityTypes childEntityType, object ordering)
        {
            Identifier = new Identifier(null,code);
            Name = name;
            ChildEntityType = childEntityType;
            Ordering = ordering;
            Parent = parent;
        }

        public GroupingEntity Parent { get; private set; }

        public object Ordering { get; private set; }

        public Identifier Identifier { get; private set;}

        public string Name { get; set; }

        public GroupingEntity Previous { get; set; }

        public Row TotalRow { get; set; }

        public RowType RowType
        {
            get
            {
                if (TotalRow!= null)
                {
                    return TotalRow.RowType;
                }
                else
                {
                    return RowTypeForNewRow;
                }
            }
        }

        protected abstract RowType RowTypeForNewRow { get;}

        public int RowNumber => TotalRow.RowNumber;

        public override string ToString()
        {
            return $"{Identifier.Code}({Name})";
        }

        public EntityTypes ChildEntityType { get; set; }

        public bool ChildrenAreDeleteable { get; set; } = false;

        public bool ChildrenAreHidden { get; set; } = false;

        public Dictionary<Identifier, IChildEntity> Children { get; set; } = new Dictionary<Identifier, IChildEntity>();

        public List<IChildEntity> ChildrenToDelete { get; set; } = new List<IChildEntity>(); 

        public string ControlString { get; set; }

        public decimal? Nav { get; set; }

        public decimal? PreviousNav { get; set; }
    }
}
