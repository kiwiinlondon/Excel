using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Identifier
    {
        public Identifier(int? id, string code)
        {
            Id = id;
            Code = code;
        }

        public int? Id { get; set; }

        public string Code { get; set; }

        protected bool Equals(Identifier other)
        {
            if (Id.HasValue && other.Id.HasValue)
            {
                return Id.Value == other.Id.Value;
            }
            else if (!string.IsNullOrWhiteSpace(Code) && !string.IsNullOrWhiteSpace(other.Code))
            {
                return string.Equals(Code.ToUpper(), other.Code.ToUpper());
            }
            return false;                    
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Identifier)obj);
        }

        public static bool operator ==(Identifier left, Identifier right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(Identifier left, Identifier right)
        {
            return !Equals(left, right);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return 1;
            }
        }

        public override string ToString()
        {
            return $"{Id}({Code})";
        }

    }
}
