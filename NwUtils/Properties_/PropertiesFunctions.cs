using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using NwUtils.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace NwUtils.Properties_
{
    public static class PropertiesFunctions
    {
        public static readonly InwOpState3 state = ComApiBridge.State;

        /// <summary>
        /// Obtém todas as propriedades de um <see cref="ModelItem"/>, organizadas por categoria.
        /// </summary>
        /// <param name="modelItem">
        /// O elemento do modelo (<see cref="ModelItem"/>) do qual as propriedades serão extraídas.
        /// </param>
        /// <param name="condition">
        /// (Opcional) Uma função de filtro (<see cref="Func{Property, Boolean}"/>) que permite
        /// selecionar apenas as propriedades que atendem a uma determinada condição.
        /// Se não for especificada, todas as propriedades do <paramref name="modelItem"/> serão retornadas.
        /// </param>
        /// <returns>
        /// Um dicionário onde a chave é o nome da categoria (<c>string</c>) e o valor é uma coleção
        /// de propriedades (<see cref="Property"/>) pertencentes a essa categoria.
        /// </returns>
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

        /// <summary>
        /// Define um conjunto de propriedades personalizadas em um <see cref="ModelItem"/> específico.
        /// </summary>
        /// <param name="target">
        /// O elemento do modelo (<see cref="ModelItem"/>) que receberá as propriedades.
        /// </param>
        /// <param name="targetCategory">
        /// O nome da categoria onde as propriedades serão criadas ou atualizadas.
        /// </param>
        /// <param name="values">
        /// Dicionário contendo os pares <c>Nome</c> e <c>Valor</c> das propriedades 
        /// que serão atribuídas ao <paramref name="target"/>.
        /// </param>
        /// <param name="forceUpdate">
        /// Indica se a interface do usuário (UI) deve ser atualizada imediatamente
        /// após a aplicação das propriedades.
        /// </param>
        public static void SetPropertiesToTarget(ModelItem target, string targetCategory, Dictionary<string, string> values, bool forceUpdate = false)
        {
            SetPropertiesToTarget(new[] { target }, targetCategory, values, forceUpdate);
        }

        /// <summary>
        /// Define um conjunto de propriedades personalizadas em uma coleção de itens do modelo (<see cref="ModelItem"/>).
        /// </summary>
        /// <param name="targets">
        /// A coleção de elementos do modelo (<see cref="ModelItem"/>) que receberão as propriedades.
        /// </param>
        /// <param name="targetCategory">
        /// O nome da categoria onde as propriedades serão criadas ou atualizadas em cada item da coleção.
        /// </param>
        /// <param name="values">
        /// Dicionário contendo os pares <c>Nome</c> e <c>Valor</c> das propriedades 
        /// que serão atribuídas a cada item do conjunto <paramref name="targets"/>.
        /// </param>
        /// <param name="forceUpdate">
        /// Indica se a interface do usuário (UI) deve ser atualizada imediatamente
        /// após a aplicação das propriedades.
        /// </param>
        /// <remarks>
        /// Esse método aplica as mesmas propriedades personalizadas a múltiplos itens do modelo.
        /// É útil em rotinas de automação, enriquecimento de dados em lote ou sincronização de atributos
        /// dentro de projetos Navisworks.
        /// </remarks>
        public static void SetPropertiesToTarget(IEnumerable<ModelItem> targets, string targetCategory, Dictionary<string, string> values, bool forceUpdate = false)
        {
            if (values.Count == 0) return;

            foreach (ModelItem model in targets)
            {
                InwOaPath3 oaPath = (InwOaPath3)ComApiBridge.ToInwOaPath(model);
                InwGUIPropertyNode2 propVec = (InwGUIPropertyNode2)state.GetGUIPropertyNode(oaPath, true);

                values = MergeProperties(values, targetCategory, model);
                InwOaPropertyVec newPropVec = (InwOaPropertyVec)state.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                foreach (var kvp in values)
                {
                    InwOaProperty newProp = (InwOaProperty)state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
                    newProp.name = kvp.Key;
                    newProp.UserName = kvp.Key;
                    newProp.value = kvp.Value;
                    newPropVec.Properties().Add(newProp);
                }

                bool categoryExists = false;
                foreach (InwGUIAttribute2 attribute in propVec.GUIAttributes())
                {
                    try
                    {
                        if (attribute.UserDefined && attribute.ClassUserName.Equals(targetCategory))
                        {
                            categoryExists = true;
                            break;
                        }
                    }
                    catch { }
                }

                if (!categoryExists)
                    propVec.SetUserDefined(0, targetCategory, targetCategory, newPropVec);

                int index = 0;
                foreach (InwGUIAttribute2 attribute in propVec.GUIAttributes())
                {
                    try
                    {
                        if (attribute.UserDefined && attribute.ClassUserName.Equals(targetCategory))
                        {
                            propVec.RemoveUserDefined(index);
                            propVec.SetUserDefined(index, targetCategory, targetCategory, newPropVec);
                        }
                        index++;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }

            if (forceUpdate)
                System.Windows.Forms.Application.DoEvents();
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
        private static Dictionary<string, string> MergeProperties(Dictionary<string, string> input, string cat, ModelItem source)
        {
            var current = GetProperties(source);
            var ofCat = current.Where(p => p.Key == cat);

            var valuesDict = new Dictionary<string, string>();
            foreach (var kvp in ofCat)
            {
                foreach (var prop in kvp.Value)
                {
                    if (prop.Name != null)
                        valuesDict[prop.Name] = prop.Value?.ToString();
                }
            }

            if (valuesDict.Count == 0)
                return input;

            foreach (var kvp in input)
            {
                if (valuesDict.ContainsKey(kvp.Key))
                {
                    valuesDict[kvp.Key] = kvp.Value;
                }
            }

            return valuesDict;
        }
    }
}
