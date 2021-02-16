using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelColumMapBuilder<T> where T : class
    {
        private static readonly ICollection<ExcelColumnConfiguration> _maps = new HashSet<ExcelColumnConfiguration>();
        public ExcelColumMapBuilder()
        {
            var properties = typeof(T).GetProperties()
                .Where(x => x.PropertyType.Assembly == typeof(object).Assembly)
                .ToArray();

            for (int i = 0; i < properties.Length; i++)
            {
                var prop = properties[i];
                _maps.Add(new ExcelColumnConfiguration
                {
                    Property = prop,
                    ColumnName = prop.Name
                });
            }
        }

        public void For<TProperty>(Expression<Func<T, TProperty>> expression, int columnIndex, string columnName = null)
        {
            var propertyName = ((MemberExpression)expression.Body).Member.Name;
            var propertyMatch = _maps.FirstOrDefault(x => x.ColumnName == propertyName);
            if (propertyMatch != null)
            {
                propertyMatch.ColumnIndex = columnIndex;
                if (!string.IsNullOrEmpty(columnName))
                    propertyMatch.ColumnName = columnName;
            }
        }

        public IReadOnlyCollection<ExcelColumnConfiguration> GetIndexMap()
        {
            return _maps.ToList().AsReadOnly();
        }
    }


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
