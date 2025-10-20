using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using NwUtils.Types;
using System;
using System.Collections.Generic;

namespace NwUtils.Properties_
{
    public static class PropertiesFunctions
    {
        public static readonly InwOpState3 state = ComApiBridge.State;

        /// <summary>
        /// Creates a new category of properties on a ModelItem
        /// Enable forceUpdate to refresh the UI immediately to reflect the changes
        /// </summary>
        public static bool SetProperty(ModelItem model, string cat, IEnumerable<Property> properties, bool forceUpdate = false)
        {
            InwOaPath oaPath = ComApiBridge.ToInwOaPath(model);
            InwGUIPropertyNode2 propVec = (InwGUIPropertyNode2)state.GetGUIPropertyNode(oaPath, true);

            InwOaPropertyVec newPropVec = (InwOaPropertyVec)state.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

            foreach (var prop in properties)
            {
                InwOaProperty newProp = (InwOaProperty)state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
                newProp.name = prop.Name;
                newProp.UserName = prop.Name;
                newProp.value = prop.DisplayValue ?? prop.Value;
                newPropVec.Properties().Add(newProp);
            }

            try
            {
                propVec.SetUserDefined(0, cat, cat, newPropVec);
                GC.KeepAlive(propVec);

                if (forceUpdate)
                    System.Windows.Forms.Application.DoEvents();

                return true;
            }
            catch 
            {
                return false;
            }
        }

        /// <summary>
        /// Gets all properties of a ModelItem, optionally filtered by a condition
        /// </summary>
        public static Dictionary<string, IEnumerable<Property>> GetProperties(ModelItem modelItem, Func<Property, bool> condition = null)
        {
            Dictionary<string, IEnumerable<Property>> result = new Dictionary<string, IEnumerable<Property>>();

            foreach (var propCategory in modelItem.PropertyCategories)
            {
                List<Property> props = new List<Property>();
                foreach (DataProperty dataProp in propCategory.Properties)
                {
                    Property prop = new Property(dataProp);

                    if (condition == null)
                        props.Add(prop);

                    else if (condition(prop))
                        props.Add(prop);
                }
                if (props.Count > 0)
                    result[propCategory.DisplayName] = props;
            }

            return result;
        }

        public static bool TryGetProperty(ModelItem modelItem, string Name, out Property property)
        {
            if (!HasProperty(modelItem, Name))
            {
                property = null;
                return false;
            }

            foreach (var propCategory in modelItem.PropertyCategories)
            {
                foreach (DataProperty dataProp in propCategory.Properties)
                {
                    if (dataProp.DisplayName.Equals(Name, StringComparison.OrdinalIgnoreCase))
                    {
                        property = new Property(dataProp);
                        return true;
                    }
                }
            }
            property = null;
            return false;
        }

        public static bool HasProperty(ModelItem modelItem, string Name)
        {
            foreach (var propCategory in modelItem.PropertyCategories)
            {
                foreach (DataProperty dataProp in propCategory.Properties)
                {
                    if (dataProp.DisplayName.Equals(Name, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }
    }
}
