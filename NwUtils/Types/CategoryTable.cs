using System;
using System.Collections.Generic;
using System.Linq;

namespace NwUtils.Types
{
    public class CategoryFrame
    {
        private Dictionary<string, IEnumerable<Property>> _values = new Dictionary<string, IEnumerable<Property>>();
        public string[] Columns => _values.Keys.ToArray();
        public int RowCount
        {
            get
            {
                if (_values.Count == 0)
                    return 0;
                return _values.First().Value.Count();
            }
        }

        public CategoryFrame(Dictionary<string, IEnumerable<Property>> values)
        {
            _values = values;
        }

        public Property this[int row, string category]
        {
            get
            {
                if (!_values.ContainsKey(category))
                    throw new ArgumentException($"Column '{category}' does not exist in the table.");
                var columnValues = _values[category].ToList();
                if (row < 0 || row >= columnValues.Count)
                    throw new IndexOutOfRangeException($"Row index '{row}' is out of range for column '{category}'.");
                return columnValues[row];
            }
        }

        public Property[] this[string category, string name]
        {
            get
            {
                if (!_values.ContainsKey(category))
                    throw new ArgumentException($"Column '{category}' does not exist in the table.");
                if (string.IsNullOrEmpty(name))
                    throw new ArgumentException("Property name cannot be null or empty.");
                if (!_values[category].Any(p => p.Name == name))
                    throw new ArgumentException($"Property '{name}' does not exist in column '{category}'.");
                return _values[category].Where(p => p.Name == name).ToArray();
            }
        }

        public Property[] this[string category]
        {
            get
            {
                if (!_values.ContainsKey(category))
                    throw new ArgumentException($"Column '{category}' does not exist in the table.");
                return _values[category].ToArray();
            }
        }

        public Property[] this[int row]
        {
            get
            {
                if (row < 0 || row >= RowCount)
                    throw new IndexOutOfRangeException($"Row index '{row}' is out of range.");
                List<Property> result = new List<Property>();
                foreach (var category in Columns)
                    result.Add(this[row, category]);
                return result.ToArray();
            }
        }

        public void AddColumn(string category, IEnumerable<Property> properties)
        {
            if (_values.ContainsKey(category))
                throw new ArgumentException($"Column '{category}' already exists in the table.");
            _values[category] = properties;
        }

    }
}
