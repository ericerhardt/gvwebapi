using System;
using System.Runtime.Serialization;

namespace GV.Lookup
{
    public class LookupInfo : IComparable
    {
        public LookupInfo(int id, string display)
        {
            Display = display;
            Id = id;
        }

        public LookupInfo()
        {
        }

        public long Id { get; set; }
        public string Display { get; set; }

        protected bool Equals(LookupInfo other)
        {
            return Id == other.Id;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((LookupInfo) obj);
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }

        public override string ToString()
        {
            return Display;
        }

        public int CompareTo(object obj)
        {
            var o = obj as LookupInfo;
            if (o != null)
                return Id.CompareTo(o.Id);
            return -1;
        }
    }
}