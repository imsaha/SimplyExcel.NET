using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace SimplyExcel.NET.Builder
{
    public class ExcelColumMapBuilder<T, TConfig> where T : class where TConfig : ExcelColumnConfiguration
    {
        private static readonly ICollection<TConfig> _maps = new HashSet<TConfig>();
        public ExcelColumMapBuilder()
        {
            var properties = typeof(T).GetProperties()
                //.Where(x => x.PropertyType.Assembly == typeof(object).Assembly)
                .ToArray();
            var currentColumnIndex = 0;
            for (int i = 0; i < properties.Length; i++)
            {
                var prop = properties[i];
                var instance = Activator.CreateInstance<TConfig>();
                instance.ColumnIndex = currentColumnIndex;
                instance.Property = prop;
                instance.ColumnName = prop.Name;
                _maps.Add(instance);

                if (prop.PropertyType.IsClass && prop.PropertyType.Assembly != typeof(object).Assembly)
                {
                    var children = new List<TConfig>();
                    var complexProperties = prop.PropertyType.GetProperties();
                    for (int j = 0; j < complexProperties.Length; j++)
                    {
                        var complexProp = complexProperties[j];
                        var child = Activator.CreateInstance<TConfig>();
                        child.ColumnIndex = j;
                        child.Property = complexProp;
                        child.ColumnName = complexProp.Name;
                        children.Add(child);

                        currentColumnIndex++;
                    }
                    instance.Children = children.ToArray();
                }
                else
                {
                    currentColumnIndex++;
                }
            }

            //var complexProperties = typeof(T).GetProperties()
            //   .Where(x => x.PropertyType.Assembly != typeof(object).Assembly)
            //   .ToArray();

            //if (complexProperties != null && complexProperties.Length > 0)
            //{
            //    var children = new List<TConfig>();
            //    for (int i = 0; i < complexProperties.Length; i++)
            //    {
            //        var prop = complexProperties[i];
            //        var instance = _maps.First(x => x.ColumnName == prop.Name);
            //        var childProperties = instance.Property.PropertyType.GetProperties();
            //        for (int j = 0; j < childProperties.Length; j++)
            //        {
            //            var chilProp = childProperties[j];
            //            var child = Activator.CreateInstance<TConfig>();
            //            child.ColumnIndex = j;
            //            child.Property = chilProp;
            //            child.ColumnName = chilProp.Name;
            //            children.Add(child);
            //        }
            //        instance.Children = children.ToArray();
            //    }
            //}

        }

        public TConfig For<TProperty>(Expression<Func<T, TProperty>> expression, int columnIndex, string columnName = null) where TProperty : IEquatable<TProperty>
        {
            var body = ((MemberExpression)expression.Body);
            var propertyName = ((MemberExpression)expression.Body).Member.Name;
            TConfig propertyMatch = _maps.FirstOrDefault(x => x.Property.Name == propertyName);
            if (body.Expression is MemberExpression)
            {
                var parentName = ((MemberExpression)body.Expression).Member.Name;
                propertyMatch = (TConfig)_maps.First(x => x.Property.Name == parentName).Children.First(x => x.Property.Name == propertyName);
            }

            if (!string.IsNullOrEmpty(columnName))
                propertyMatch.ColumnName = columnName;

            //if (columnIndex > -1)
            //{
            //    for (int i = 0; i < _maps.Count; i++)
            //    {
            //        var nextItem = _maps.ElementAtOrDefault(i + 1);
            //        var item = _maps.ElementAt(i);
            //        item.ColumnIndex = (i - 1) + columnIndex;
            //        if (item.ColumnIndex == nextItem?.ColumnIndex)
            //        {
            //            nextItem.ColumnIndex = +1;
            //        }

            //        //if (item.Children != null && item.Children.Length > 0)
            //        //{
            //        //    for (int j = 0; j < item.Children.Length; j++)
            //        //    {
            //        //        var nextChildItem = item.Children.ElementAtOrDefault(j + 1);
            //        //        var childItem = item.Children.ElementAt(j);
            //        //        childItem.ColumnIndex = (j - 1) + columnIndex;
            //        //        if (childItem.ColumnIndex == nextChildItem?.ColumnIndex)
            //        //        {
            //        //            nextChildItem.ColumnIndex = +1;
            //        //        }
            //        //    }
            //        //}
            //    }
            //}

            return propertyMatch;
        }

        public IReadOnlyCollection<TConfig> GetMaps()
        {
            return _maps.ToList().AsReadOnly();
        }
    }
}
