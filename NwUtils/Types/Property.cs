using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NwUtils.Types
{
    /// <summary>
    /// Property class that represents a DataProperty from Navisworks
    /// </summary>
    public class Property
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public string DisplayValue { get; set; }

        public Property(string name, object value, string displayValue = null)
        {
            Name = name;
            Value = value;
            DisplayValue = displayValue ?? value?.ToString();
        }

        public Property(DataProperty dataProperty)
        {
            Name = dataProperty.DisplayName;
            Value = GetPropertyValue(dataProperty);
            DisplayValue = dataProperty.Value.IsDisplayString ? dataProperty.Value.ToDisplayString() : Value?.ToString();
        }

        public override string ToString()
        {
            return $"{Name}: {DisplayValue}";
        }

        /// <summary>
        /// Retrieves the value of a DataProperty in its appropriate .NET type
        /// </summary>
        public static object GetPropertyValue(DataProperty property)
        {
            if (property == null || property.Value == null)
                return null;

            var val = property.Value;

            switch (val.DataType)
            {
                case VariantDataType.None:
                    return null;

                case VariantDataType.Int32:;
                    return val.ToInt32();

                case VariantDataType.Double:
                    return val.ToDouble();

                case VariantDataType.DoubleLength:
                    return val.ToDoubleLength();

                case VariantDataType.DoubleAngle:
                    return val.ToDoubleAngle();

                case VariantDataType.DoubleArea:
                    return val.ToDoubleArea();

                case VariantDataType.DoubleVolume:
                    return val.ToDoubleVolume();

                case VariantDataType.Boolean:
                    return val.ToBoolean();

                case VariantDataType.DisplayString:
                    return val.ToDisplayString();

                case VariantDataType.DateTime:
                    return val.ToDateTime();

                case VariantDataType.NamedConstant:
                    return val.ToNamedConstant();

                case VariantDataType.IdentifierString:;
                    return val.ToString();

                case VariantDataType.Point2D:
                    return val.ToPoint2D();

                case VariantDataType.Point3D:
                    return val.ToPoint3D();

                default:
                    return val.ToString();
            }
        }
    }
}
