
using System.Diagnostics;
namespace McKinsey.PowerPointGenerator.Core.Data
{
    [DebuggerDisplay("{Name} | {Number}, hidden: {IsHidden}")]
    public class Index
    {
        public string Name { get; set; }
        public int? Number { get; set; }
        public bool IsAll { get; set; }
        public bool IsHidden { get; set; }
        public bool IsCore { get; set; }

        public Index(int index)
        {
            Number = index;
            IsAll = false;
        }

        public Index(string value)
        {   
            if (value.Trim().Trim('"') == "*")
            {
                IsAll = true;
                return;
            }
            IsAll = false;
            int idx;
            if (int.TryParse(value.Trim('"'), out idx))
            {
                Number = idx;
            }
            else
            {
                Name = value.Trim('"');
            }
        }

        public static bool operator ==(Index a, Index b)
        {
            if (System.Object.ReferenceEquals(a, b))
            {
                return true;
            }
            if (((object)a == null) || ((object)b == null))
            {
                return false;
            }
            if (a.Number.HasValue && b.Number.HasValue && a.Number.Value == b.Number.Value)
            {
                return true;
            }
            if (a.IsAll && b.IsAll)
            {
                return true;
            }
            if (!string.IsNullOrEmpty(a.Name) && !string.IsNullOrEmpty(b.Name) && a.Name.Equals(b.Name, System.StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public static bool operator !=(Index a, Index b)
        {
            return !(a == b);
        }
    }
}
