using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;

namespace GVWebapi.Helpers
{
    public static class PiviotArrays
    {
        public static dynamic[] ToPivotArray<T, TColumn, TGroupRow, TRow, TData>(
        this IEnumerable<T> source,
        Func<T, TColumn> columnSelector,
        Expression<Func<T, TGroupRow>> rowSelector,
        Expression<Func<T, TRow>> extraSelector,
        Func<IEnumerable<T>, TData> dataSelector)
        {
            if (rowSelector == null)
            {
                throw new ArgumentNullException(nameof(rowSelector));
            }
            if (extraSelector == null)
            {
                throw new ArgumentNullException(nameof(extraSelector));
            }

            var arr = new List<object>();
            var cols = new List<string>();
            String rowName = ((MemberExpression)rowSelector.Body).Member.Name;
            String extraCol = ((MemberExpression)extraSelector.Body).Member.Name;
            var columns = source.Select(columnSelector).Distinct();

            cols = (new[] { extraCol }).Concat(new[] { rowName }).Concat(columns.Select(x => x.ToString())).ToList();

            var rows = source.GroupBy(rowSelector.Compile())
                             .Select(rowGroup => new
                             {
                                CostCenter = rowGroup.Key,                             
                                Values = columns.GroupJoin(
                                rowGroup,
                                c => c,
                                s => columnSelector(s),
                                (c, columnGroup) => dataSelector(columnGroup))
                             }).ToArray();


            foreach (var row in rows)
            {
                var items = row.Values.Cast<object>().ToList();
                items.Insert(0, row.CostCenter);
                items.Insert(0, 0);
                var obj = GetAnonymousObject(cols, items);
                arr.Add(obj);
            }
            return arr.ToArray();
        }
        private static dynamic GetAnonymousObject(IEnumerable<string> columns, IEnumerable<object> values)
        {
            IDictionary<string, object> eo = new ExpandoObject() as IDictionary<string, object>;
            int i;
            for (i = 0; i < columns.Count(); i++)
            {
                eo.Add(columns.ElementAt<string>(i), values.ElementAt<object>(i));
            }
            return eo;
        }
    }
         
}