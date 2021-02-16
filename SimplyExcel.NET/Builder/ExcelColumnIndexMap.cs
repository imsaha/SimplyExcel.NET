using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelColumnIndexMap<T>
    {
        private static readonly Dictionary<PropertyInfo, int> _matches = new Dictionary<PropertyInfo, int>();
        public ExcelColumnIndexMap()
        {
            var properties = typeof(T).GetProperties()
                .Where(x => x.PropertyType.Assembly == typeof(object).Assembly)
                .ToArray();

            for (int i = 0; i < properties.Length; i++)
            {
                _matches.Add(properties[i], i);
            }
        }

        public Dictionary<PropertyInfo, int> GetIndexMap()
        {
            return _matches;
        }

        public ExcelColumnIndexMap<T> For<TProperty>(Expression<Func<T, TProperty>> expression, int columnIndex)
        {
            var propertyName = ((MemberExpression)expression.Body).Member.Name;
            var propertyMatch = _matches.FirstOrDefault(x => x.Key.Name == propertyName);
            _matches.Remove(propertyMatch.Key);
            _matches.Add(propertyMatch.Key, columnIndex);
            return this;
        }
    }
}
