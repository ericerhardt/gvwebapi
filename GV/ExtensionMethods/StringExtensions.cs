using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace GV.ExtensionMethods
{
    public static class StringExtensions
    {
        public static bool EqualsIgnore(this string source, string value)
        {
            return source.Equals(value, StringComparison.OrdinalIgnoreCase);
        }

        public static IEnumerable<List<T>> ChunkBy<T>(this List<T> locations, int chunkSize)
        {
            for (var i = 0; i < locations.Count; i += chunkSize)
            {
                yield return locations.GetRange(i, Math.Min(chunkSize, locations.Count - i));
            }
        }

        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            var properties = TypeDescriptor.GetProperties(typeof(T));
            var table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (var item in data)
            {
                var row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }
    }
}