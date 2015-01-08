using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using QUT.Xbim.Gppg;
using Xbim.IO;
using Xbim.XbimExtensions.Interfaces;
using Xbim.Ifc2x3.Kernel;
using System.Linq.Expressions;
using System.Reflection;
using Xbim.XbimExtensions.SelectTypes;
using Xbim.Ifc2x3.MaterialResource;
using Xbim.Ifc2x3.Extensions;
using Xbim.Ifc2x3.MeasureResource;
using Xbim.Ifc2x3.ProductExtension;
using System.IO;
using Xbim.Ifc2x3.PropertyResource;
using Xbim.Ifc2x3.MaterialPropertyResource;
using Xbim.Ifc2x3.QuantityResource;
using Xbim.Ifc2x3.ExternalReferenceResource;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Globalization;
using Xbim.COBie;
using Xbim.COBie.Serialisers;
using Newtonsoft.Json;

namespace Xbim.Script
{
    internal partial class Parser
    {
        private XbimModel _model;
        private XbimVariables _variables;
        private ParameterExpression _input = Expression.Parameter(typeof(IPersistIfcEntity), "Input");

        //public properties of the parser
        public XbimVariables Variables { get { return _variables; } }
        public XbimModel Model { get { return _model; } }
        public TextWriter Output { get; set; }

        internal Parser(Scanner lex, XbimModel model): base(lex)
        {
            _model = model;
            _variables = new XbimVariables();
            if (_model == null) throw new ArgumentNullException("Model is NULL");
        }

        #region Objects creation
        private IPersistIfcEntity CreateObject(Type type, string name, string description = null)
        {
            if (_model == null) throw new ArgumentNullException("Model is NULL");
            if (name == null)
            {
                Scanner.yyerror("Name must be defined for creation of the " + type.Name + ".");
            } 

            Func<IPersistIfcEntity> create = () => {
                var result = _model.Instances.New(type);

                //set name and description
                if (result == null) return null;
                IfcRoot root = result as IfcRoot;
                if (root != null)
                {
                    root.Name = name;
                    root.Description = description;
                }
                IfcMaterial material = result as IfcMaterial;
                if (material != null)
                {
                    material.Name = name;
                }

                return result;
            };

            IPersistIfcEntity entity = null;
            if (_model.IsTransacting)
            {
                entity = create();
            }
            else
            {
                using (var txn = _model.BeginTransaction("Object creation"))
                {
                    entity = create();
                    txn.Commit();
                }
            }
            return entity;
        }

        private IfcMaterialLayerSet CreateLayerSet(string name, List<Layer> layers)
        {
            Func<IfcMaterialLayerSet> create = () => {
                return _model.Instances.New<IfcMaterialLayerSet>(ms =>
                {
                    ms.LayerSetName = name;
                    foreach (var layer in layers)
                    {
                        ms.MaterialLayers.Add(_model.Instances.New<IfcMaterialLayer>(ml =>
                        {
                            ml.LayerThickness = layer.thickness;
                            //get material if it already exists
                            var material = _model.Instances.Where<IfcMaterial>(m => m.Name.ToString().ToLower() == layer.material).FirstOrDefault();
                            if (material == null)
                                material = _model.Instances.New<IfcMaterial>(m => m.Name = layer.material);
                            ml.Material = material;
                        }));
                    }
                });
            };

            IfcMaterialLayerSet result = null;
            if (_model.IsTransacting)
                result = create();
            else
                using (var txn = _model.BeginTransaction())
                {
                    result = create();
                    txn.Commit();
                }
            return result;
        }
        #endregion

        #region Attribute and property conditions
        private Expression GenerateValueCondition(string property, object value, Tokens condition, Tokens type)
        {
            var propNameExpr = Expression.Constant(property);
            var valExpr = Expression.Constant(value, typeof(object));
            var condExpr = Expression.Constant(condition);
            var thisExpr = Expression.Constant(this);

            string method = null;
            switch (type)
            {
                case Tokens.STRING:
                    method = "EvaluateValueCondition";
                    break;
                case Tokens.PROPERTY:
                    method = "EvaluatePropertyCondition";
                    break;
                case Tokens.ATTRIBUTE:
                    method = "EvaluateAttributeCondition";
                    break;
                default:
                    throw new ArgumentException("Unexpected value of the 'type'. Expected values: STRING, PROPERTY, ATTRIBUTE.");
            }

            var evaluateMethod = GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic);

            return Expression.Call(thisExpr, evaluateMethod, _input, propNameExpr, valExpr, condExpr);
        }

        private bool EvaluateValueCondition(IPersistIfcEntity input, string propertyName, object value, Tokens condition)
        {
            //try to get attribute
            var attr = GetAttributeValue(propertyName, input);
            var prop = attr as IfcValue;
            if (propertyName.ToLower() == "entitylabel")
                prop = new IfcInteger(Math.Abs((int)attr));

            //try to get property if attribute doesn't exist
            if (prop == null)
             prop = GetPropertyValue(propertyName, input);

            return EvaluateValue(prop, value, condition);
        }

        private bool EvaluatePropertyCondition(IPersistIfcEntity input, string propertyName, object value, Tokens condition)
        {
            var prop = GetPropertyValue(propertyName, input);
            return EvaluateValue(prop, value, condition);
        }

        private bool EvaluateAttributeCondition(IPersistIfcEntity input, string attribute, object value, Tokens condition)
        {
            var attr = GetAttributeValue(attribute, input);
            return EvaluateValue(attr, value, condition);
        }
        #endregion

        #region Property and attribute conditions helpers
        private static bool EvaluateNullCondition(object expected, object actual, Tokens condition)
        {
            if (expected != null && actual != null)
                throw new ArgumentException("One of the values is expected to be null.");
            switch (condition)
            {
                case Tokens.OP_EQ:
                    return expected == null && actual == null;
                case Tokens.OP_NEQ:
                    if (expected == null && actual != null) return true;
                    if (expected != null && actual == null) return true;
                    return false;
                default:
                    return false;
            }
        }

        private bool EvaluateValue(object ifcVal, object val, Tokens condition)
        {
            //special handling for null value comparison
            if (val == null || ifcVal == null)
            {
                try
                {
                    return EvaluateNullCondition(ifcVal, val, condition);
                }
                catch (Exception e)
                {
                    Scanner.yyerror(e.Message);
                    return false;
                }
            }

            //try to get values to the same level; none of the values can be null for this operation
            object left = null;
            object right = null;
            try
            {
                left = UnWrapType(ifcVal);
                right = PromoteType(GetNonNullableType(left.GetType()), val);
            }
            catch (Exception)
            {

                Scanner.yyerror(val.ToString() + " is not compatible type with type of " + ifcVal.GetType());
                return false;
            }
            

            //create expression
            bool? result = null;
            switch (condition)
            {
                case Tokens.OP_EQ:
                    if (left is string  && right is string)
                        return ((string)left).ToLower() == ((string)right).ToLower();
                    return left.Equals(right);
                case Tokens.OP_NEQ:
                    if (left is string && right is string)
                        return ((string)left).ToLower() != ((string)right).ToLower();
                    return !left.Equals(right);
                case Tokens.OP_GT:
                    result = GreaterThan(left, right);
                    if (result != null) return result ?? false;
                    break;
                case Tokens.OP_LT:
                    result = LessThan(left, right);
                    if (result != null) return result ?? false;
                    break;
                case Tokens.OP_GTE:
                    result = !LessThan(left, right);
                    if (result != null) return result ?? false;
                    break;
                case Tokens.OP_LTQ:
                    result = !GreaterThan(left, right);
                    if (result != null) return result ?? false;
                    break;
                case Tokens.OP_CONTAINS:
                    return Contains(left, right);
                case Tokens.OP_NOT_CONTAINS:
                    return !Contains(left, right);
                default:
                    throw new ArgumentOutOfRangeException("Unexpected token used as a condition");
            }
            Scanner.yyerror("Can't compare " + left + " and " + right + ".");
            return false;
        }

        private static object UnWrapType(object value)
        { 
            //enumeration
            if (value.GetType().IsEnum)
                return Enum.GetName(value.GetType(), value);

            //express type
            ExpressType express = value as ExpressType;
            if (express != null)
                return express.Value;

            return value;
        }

        private static bool IsNullableType(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        private static Type GetNonNullableType(Type type)
        {
            return IsNullableType(type) ? type.GetGenericArguments()[0] : type;
        }

        private static bool IsOfType(Type type, IPersistIfcEntity entity)
        {
            return type.IsAssignableFrom(entity.GetType());
        }

        private static PropertyInfo GetAttributeInfo(string name, IPersistIfcEntity entity)
        {
            Type type = entity.GetType();
            return type.GetProperty(name, BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.Public);
        }

        private static object GetAttributeValue(string name, IPersistIfcEntity entity)
        {
            PropertyInfo pInfo = GetAttributeInfo(name, entity);
            if (pInfo == null)
                return null;
            return pInfo.GetValue(entity, null);
        }

        private static PropertyInfo GetPropertyInfo(string name, IPersistIfcEntity entity, out object propertyObject)
        {
            //try to get the name of the pSet if it is encoded in there
            var split = name.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            string pSetName = null;
            if (split.Count() == 2)
            {
                pSetName = split[0];
                name = split[1];
            }
            var specificPSet = pSetName != null;

            List<IfcPropertySet> pSets = null;
            IEnumerable<IfcExtendedMaterialProperties> pSetsMaterial = null;
            IEnumerable<IfcElementQuantity> elQuants = null;
            IfcPropertySingleValue property = null;
            IfcPhysicalSimpleQuantity quantity = null;
            IfcPropertySet ps = null;
            IfcElementQuantity eq = null;
            IfcExtendedMaterialProperties eps = null;

            IfcObject obj = entity as IfcObject;
            if (obj != null)
            {
                if (specificPSet)
                {
                    ps = obj.GetPropertySet(pSetName);
                    eq = obj.GetElementQuantity(pSetName);
                }
                pSets =  ps == null ? obj.GetAllPropertySets() : new List<IfcPropertySet>(){ps};
                elQuants = eq == null ? obj.GetAllElementQuantities() : new List<IfcElementQuantity>() { eq };
            }
            IfcTypeObject typeObj = entity as IfcTypeObject;
            if (typeObj != null)
            {
                if (specificPSet)
                {
                    ps = typeObj.GetPropertySet(pSetName);
                    eq = typeObj.GetElementQuantity(pSetName);
                }
                pSets = ps == null ? typeObj.GetAllPropertySets() : new List<IfcPropertySet>() { ps };
                elQuants = eq == null ? typeObj.GetAllElementQuantities() : new List<IfcElementQuantity>() { eq};
            }
            IfcMaterial material = entity as IfcMaterial;
            if (material != null)
            {
                if (specificPSet)
                    eps = material.GetExtendedProperties(pSetName);
                pSetsMaterial = eps == null ? material.GetAllPropertySets() : new List<IfcExtendedMaterialProperties>() { eps };
            }

            if (pSets != null)
                foreach (var pSet in pSets)
                {
                    foreach (var prop in pSet.HasProperties)
                    {
                        if (prop.Name.ToString().ToLower() == name.ToLower()) property = prop as IfcPropertySingleValue;
                    }
                }
            if (pSetsMaterial != null)
                foreach (var pSet in pSetsMaterial)
                {
                    foreach (var prop in pSet.ExtendedProperties)
                    {
                        if (prop.Name.ToString().ToLower() == name.ToLower()) property = prop as IfcPropertySingleValue;
                    }
                }
            if (elQuants != null)
                foreach (var quant in elQuants)
                {
                    foreach (var item in quant.Quantities)
                    {
                        if (item.Name.ToString().ToLower() == name.ToLower()) quantity = item as IfcPhysicalSimpleQuantity;
                    }
                }


            //set property
            if (property != null)
            {
                propertyObject = property;
                return property.GetType().GetProperty("NominalValue");
            }

            //set simple quantity
            else if (quantity != null)
            {
                PropertyInfo info = null;
                var qType = quantity.GetType();
                switch (qType.Name)
                {
                    case "IfcQuantityLength":
                        info = qType.GetProperty("LengthValue");
                        break;
                    case "IfcQuantityArea":
                        info = qType.GetProperty("AreaValue");
                        break;
                    case "IfcQuantityVolume":
                        info = qType.GetProperty("VolumeValue");
                        break;
                    case "IfcQuantityCount":
                        info = qType.GetProperty("CountValue");
                        break;
                    case "IfcQuantityWeight":
                        info = qType.GetProperty("WeightValue");
                        break;
                    case "IfcQuantityTime":
                        info = qType.GetProperty("TimeValue");
                        break;
                    default:
                        throw new NotImplementedException();
                }
                if (info != null)
                {
                    propertyObject = quantity;
                    return info;
                }
            }

            propertyObject = null;
            return null;
        }


        private static IfcValue GetPropertyValue(string name, IPersistIfcEntity entity)
        {
            object pObject = null;
            var pInfo = GetPropertyInfo(name, entity, out pObject);

            if (pInfo == null) return null;

            var result = pInfo.GetValue(pObject, null);
            return result as IfcValue;
        }


        private static object PromoteType(Type targetType, object value)
        {

            if (targetType == typeof(Boolean)) return Convert.ToBoolean(value);
            if (targetType == typeof(Byte)) return Convert.ToByte(value);
            if (targetType == typeof(DateTime)) return Convert.ToDateTime(value);
            if (targetType == typeof(Decimal)) return Convert.ToDecimal(value);
            if (targetType == typeof(Double)) return Convert.ToDouble(value);
            if (targetType == typeof(float)) return Convert.ToDouble(value);
            if (targetType == typeof(Char)) return Convert.ToChar(value);
            if (targetType == typeof(Int16)) return Convert.ToInt16(value);
            if (targetType == typeof(Int32)) return Convert.ToInt32(value);
            if (targetType == typeof(Int64)) return Convert.ToInt64(value);
            if (targetType == typeof(SByte)) return Convert.ToSByte(value);
            if (targetType == typeof(Single)) return Convert.ToSingle(value);
            if (targetType == typeof(String)) return Convert.ToString(value);
            if (targetType == typeof(UInt16)) return Convert.ToUInt16(value);
            if (targetType == typeof(UInt32)) return Convert.ToUInt32(value);
            if (targetType == typeof(UInt64)) return Convert.ToUInt64(value);

            throw new Exception("Unexpected type");
        }

        private static bool? GreaterThan(object left, object right) 
        {
            try
            {
                var leftD = Convert.ToDouble(left);
                var rightD = Convert.ToDouble(right);
                return leftD > rightD;
            }
            catch (Exception)
            {
                return null;   
            }
           
        }

        private static bool? LessThan(object left, object right)
        {
            try
            {
                var leftD = Convert.ToDouble(left);
                var rightD = Convert.ToDouble(right);
                return leftD < rightD;
            }
            catch (Exception)
            {
                return null;
            }

        }

        private static bool Contains(object left, object right)
        {
            string leftS = Convert.ToString(left);
            string rightS = Convert.ToString(right);

            return leftS.ToLower().Contains(rightS.ToLower());
        }
        #endregion

        #region Select statements
        private IEnumerable<IPersistIfcEntity> Select(Type type, string name)
        {
            if (!typeof(IfcRoot).IsAssignableFrom(type)) return new IPersistIfcEntity[]{};
            Expression expression = GenerateValueCondition("Name", name, Tokens.OP_EQ, Tokens.ATTRIBUTE);
            return Select(type, expression);
        }

        private IEnumerable<IPersistIfcEntity> Select(Type type, Expression condition = null)
        {
            MethodInfo method = _model.Instances.GetType().GetMethod("OfType", new Type[] { typeof(bool) });
            MethodInfo generic = method.MakeGenericMethod(type);
            if (condition != null)
            {
                var typeFiltered = generic.Invoke(_model.Instances, new object[] { true }) as IEnumerable<IPersistIfcEntity>;
                return typeFiltered.Where(Expression.Lambda<Func<IPersistIfcEntity, bool>>(condition, _input).Compile());
            }
            else
            {
                var typeFiltered = generic.Invoke(_model.Instances, new object[] { false }) as IEnumerable<IPersistIfcEntity>;
                return typeFiltered;
            }

        }

        private IEnumerable<IPersistIfcEntity> SelectClassification(string code)
        {
            return _model.Instances.Where<IfcClassificationReference>(c => c.ItemReference.ToString().ToLower() == code.ToLower());
        }
        
        #endregion

        #region TypeObject conditions 
        private Expression GenerateTypeObjectNameCondition(string typeName, Tokens condition)
        {
            var typeNameExpr = Expression.Constant(typeName);
            var condExpr = Expression.Constant(condition);
            var thisExpr = Expression.Constant(this);

            var evaluateMethod = GetType().GetMethod("EvaluateTypeObjectName", BindingFlags.Instance | BindingFlags.NonPublic);
            return Expression.Call(thisExpr, evaluateMethod, _input, typeNameExpr, condExpr);
        }

        private Expression GenerateTypeObjectTypeCondition(Type type, Tokens condition)
        {
            var typeExpr = Expression.Constant(type, typeof(Type));
            var condExpr = Expression.Constant(condition);
            var thisExpr = Expression.Constant(this);

            var evaluateMethod = GetType().GetMethod("EvaluateTypeObjectType", BindingFlags.Instance | BindingFlags.NonPublic);
            return Expression.Call(thisExpr, evaluateMethod, _input, typeExpr, condExpr);
        }

        private bool EvaluateTypeObjectName(IPersistIfcEntity input, string typeName, Tokens condition)
        {
            IfcObject obj = input as IfcObject;
            if (obj == null) return false;

            var type = obj.GetDefiningType();
           
            //null variant
            if (type == null)
            {
                return false;
            }

            switch (condition)
            {
                case Tokens.OP_EQ:
                    return type.Name == typeName;
                case Tokens.OP_NEQ:
                    return type.Name != typeName;
                case Tokens.OP_CONTAINS:
                    return type.Name.ToString().ToLower().Contains(typeName.ToLower());
                case Tokens.OP_NOT_CONTAINS:
                    return !type.Name.ToString().ToLower().Contains(typeName.ToLower());
                default:
                    Scanner.yyerror("Unexpected Token in this function. Only equality or containment expected.");
                    return false;
            }
        }

        private bool EvaluateTypeObjectType(IPersistIfcEntity input, Type type, Tokens condition)
        {
            IfcObject obj = input as IfcObject;
            if (obj == null) return false;

            var typeObj = obj.GetDefiningType();
            
            //null variant
            if (typeObj == null || type == null)
            {
                try
                {
                    return EvaluateNullCondition(typeObj, type, condition);
                }
                catch (Exception e)
                {
                    Scanner.yyerror(e.Message);
                    return false;
                }
            }

            switch (condition)
            {
                case Tokens.OP_EQ:
                    return typeObj.GetType() == type;
                case Tokens.OP_NEQ:
                    return typeObj.GetType() != type;
                default:
                    Scanner.yyerror("Unexpected Token in this function. Only OP_EQ or OP_NEQ expected.");
                    return false;
            }
        }

        private Expression GenerateTypeCondition(Expression expression) 
        {
            var function = Expression.Lambda<Func<IPersistIfcEntity, bool>>(expression, _input).Compile();
            var fceExpr = Expression.Constant(function);
            var thisExpr = Expression.Constant(this);

            var evaluateMethod = GetType().GetMethod("EvaluateTypeCondition", BindingFlags.Instance | BindingFlags.NonPublic);

            return Expression.Call(thisExpr, evaluateMethod, _input, fceExpr);
        }

        private bool EvaluateTypeCondition(IPersistIfcEntity input, Func<IPersistIfcEntity, bool> function)
        {
            var obj = input as IfcObject;
            if (obj == null) return false;

            var defObj = obj.GetDefiningType();
            if (defObj == null) return false;

            return function(defObj);
        }

        #endregion

        #region Classification conditions

        private Expression GenerateClassificationCondition(string code, Tokens condition)
        {
            var codeExpr = Expression.Constant(code, typeof(String));
            var condExpr = Expression.Constant(condition);
            var thisExpr = Expression.Constant(this);

            var evaluateMethod = GetType().GetMethod("EvaluateClassificationCondition", BindingFlags.Instance | BindingFlags.NonPublic);

            return Expression.Call(thisExpr, evaluateMethod, _input, codeExpr, condExpr);
        }

        private bool EvaluateClassificationCondition(IPersistIfcEntity input, string code, Tokens condition)
        {
            var root = input as IfcRoot;
            if (root == null)
            {
                Scanner.yyerror("Object of type {0} can't have a classification associated.", input.GetType().Name);
                return false;
            }

            var codes = new List<string>();
            var model = input.ModelOf;
            var rels = model.Instances.Where<IfcRelAssociatesClassification>(r => r.RelatedObjects.Contains(input));
            foreach (var rel in rels)
            {
                string classCode = null;
                var reference = rel.RelatingClassification as IfcClassificationReference;
                if (reference != null) 
                    classCode = reference.ItemReference;
                
                var notation = rel.RelatingClassification as IfcClassificationNotation;
                if (notation != null)
                    classCode = ConcatClassFacets(notation.NotationFacets);

                if (!String.IsNullOrEmpty( classCode))
                {
                    codes.Add(classCode.ToLower());
                }
            }

            if (code == null)
            {
                switch (condition)
                {
                    case Tokens.OP_EQ:
                        return codes.Count == 0; 
                    case Tokens.OP_NEQ:
                        return codes.Count != 0;
                    default:
                        throw new ArgumentException("Unexpected value of the 'condition' variable. OP_EQ or OP_NEQ expected only.");
                }
            }
            else
            {
                code = code.ToLower(); //normalization for case insensitive search
                switch (condition)
                {
                    case Tokens.OP_EQ:
                        return codes.Contains(code);
                    case Tokens.OP_NEQ:
                        return !codes.Contains(code);
                    default:
                        throw new ArgumentException("Unexpected value of the 'condition' variable. OP_EQ or OP_NEQ expected only.");
                }
            }

        }

        private string ConcatClassFacets(IEnumerable<IfcClassificationNotationFacet> facets)
        {
            var result = "";
            foreach (var item in facets)
            {
                result += item.NotationValue + "-";
            }
            //trim from the last '-'
            result = result.Trim('-');
            return result;
        }
        
        #endregion

        #region Material conditions
        private Expression GenerateMaterialCondition(string materialName, Tokens condition)
        {
            Expression nameExpr = Expression.Constant(materialName);
            Expression condExpr = Expression.Constant(condition);

            var evaluateMethod = GetType().GetMethod("EvaluateMaterialCondition", BindingFlags.Static | BindingFlags.NonPublic);
            return Expression.Call(null, evaluateMethod, _input, nameExpr, condExpr);
        }

        private static bool EvaluateMaterialCondition(IPersistIfcEntity input, string materialName, Tokens condition) 
        {
            IfcRoot root = input as IfcRoot;
            if (root == null) return false;
            IModel model = root.ModelOf;

            var materialRelations = model.Instances.Where<IfcRelAssociatesMaterial>(r => r.RelatedObjects.Contains(root));
            List<string> names = new List<string>();
            foreach (var mRel in materialRelations)
            {
                names.AddRange(GetMaterialNames(mRel.RelatingMaterial));    
            }

            //convert to lower case
            for (int i = 0; i < names.Count; i++)
                names[i] = names[i].ToLower();

            switch (condition)
            {

                case Tokens.OP_EQ:
                    return names.Contains(materialName.ToLower());
                case Tokens.OP_NEQ:
                    return !names.Contains(materialName.ToLower());
                case Tokens.OP_CONTAINS:
                    foreach (var name in names)
                    {
                        if (name.Contains(materialName.ToLower())) return true;
                    }
                    break;
                case Tokens.OP_NOT_CONTAINS:
                    foreach (var name in names)
                    {
                        if (name.Contains(materialName.ToLower())) return false;
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException("Unexpected Token value.");
            }
            return false;
        }

        /// <summary>
        /// Get names of all materials involved
        /// </summary>
        /// <param name="materialSelect">Possible types of material</param>
        /// <returns>List of names</returns>
        private static List<string> GetMaterialNames(IfcMaterialSelect materialSelect)
        {
            List<string> names = new List<string>();
            
            IfcMaterial material = materialSelect as IfcMaterial;
            if (material != null) names.Add( material.Name);

            IfcMaterialList materialList = materialSelect as IfcMaterialList;
            if (materialList != null)
                foreach (var m in materialList.Materials)
                {
                    names.Add(m.Name);
                }
            
            IfcMaterialLayerSetUsage materialUsage = materialSelect as IfcMaterialLayerSetUsage;
            if (materialUsage != null)
                names.AddRange(GetMaterialNames(materialUsage.ForLayerSet));
            
            IfcMaterialLayerSet materialLayerSet = materialSelect as IfcMaterialLayerSet;
            if (materialLayerSet != null)
                foreach (var m in materialLayerSet.MaterialLayers)
                {
                    names.AddRange(GetMaterialNames(m));
                }
            
            IfcMaterialLayer materialLayer = materialSelect as IfcMaterialLayer;
            if (materialLayer != null)
                if (materialLayer.Material != null)
                    names.Add(materialLayer.Material.Name);

            return names;
        }
        #endregion

        #region Model conditions
        private Expression GenerateModelCondition(Tokens type, Tokens condition, string value)
        {
            var valExpr = Expression.Constant(value);
            var typeExpr = Expression.Constant(type);
            var condExpr = Expression.Constant(condition);
            var thisExpr = Expression.Constant(this);
            var modelExpr = Expression.Constant(_model);

            var evaluateMethod = GetType().GetMethod("EvaluateModelCondition", BindingFlags.Instance | BindingFlags.NonPublic);
            return Expression.Call(thisExpr, evaluateMethod, _input, valExpr, typeExpr, condExpr);
        }

        private bool EvaluateModelCondition(IPersistIfcEntity input, string value, Tokens type, Tokens condition)
        {
            IModel model = input.ModelOf;
            IModel testModel = null;

            foreach (var refMod in _model.ReferencedModels)
            {
                switch (type)
                {
                    case Tokens.MODEL:
                        if (IsNameOfModel(refMod.Name, value))
                            testModel = refMod.Model;
                        break;
                    case Tokens.OWNER:
                        if (refMod.OwnerName.ToLower() == value.ToLower())
                            testModel = refMod.Model;
                        break;
                    case Tokens.ORGANIZATION:
                        if (refMod.OrganisationName.ToLower() == value.ToLower())
                            testModel = refMod.Model;
                        break;

                    default:
                        throw new ArgumentException("Unexpected condition. Only MODEL, OWNER or ORGANIZATION expected.");
                }
                if (testModel != null) 
                    break;
            }

            switch (condition)
            {
                case Tokens.OP_EQ:
                    return model == testModel;
                case Tokens.OP_NEQ:
                    return model != testModel;
                default:
                    throw new ArgumentException("Unexpected condition. Only OP_EQ or OP_NEQ expected.");
            }
        }

        private static bool IsNameOfModel(string modelName, string name)
        {
            var mName = modelName.ToLower();
            var sName = name.ToLower();

            if (mName == sName) return true;
            if (Path.GetFileName(mName) == sName) return true;
            if (Path.GetFileNameWithoutExtension(mName) == sName) return true;

            return false;
        }
        #endregion

        #region Variables manipulation
        private void AddOrRemoveFromSelection(string variableName, Tokens operation, object entities)
        {
            IEnumerable<IPersistIfcEntity> ent = entities as IEnumerable<IPersistIfcEntity>;
            if (ent == null) throw new ArgumentException("Entities should be IEnumerable<IPersistIfcEntity>");
            switch (operation)
            {
                case Tokens.OP_EQ:
                    _variables.AddEntities(variableName, ent);
                    break;
                case Tokens.OP_NEQ:
                    _variables.RemoveEntities(variableName, ent);
                    break;
                default:
                    throw new ArgumentException("Unexpected token. OP_EQ or OP_NEQ expected only.");
            }
        }

        private IEnumerable<IPersistIfcEntity> GetVariableContent(string identifier)
        {
            if (string.IsNullOrEmpty(identifier))
                throw new ArgumentNullException();

            if (_variables.IsDefined(identifier))
                return _variables[identifier];
            else
            {
                Scanner.yyerror("Identifier {0} is not defined", identifier);
                return new IPersistIfcEntity[] { };
            }
        }

        private void DumpIdentifier(string identifier, string outputPath = null)
        {
            TextWriter output = null;
            if (outputPath != null)
            {
                output = new StreamWriter(outputPath, false);
            }

            StringBuilder str = new StringBuilder();
            if (Variables.IsDefined(identifier))
            {
                foreach (var entity in Variables[identifier])
                {
                    if (entity != null)
                    {
                        var name = GetAttributeValue("Name", entity);
                        str.AppendLine(String.Format("{1} #{0}: {2}", entity.EntityLabel, entity.GetType().Name, name != null ? name.ToString() : "No name defined"));
                    }
                    else
                        throw new Exception("Null entity in the dictionary");
                }
            }
            else
                str.AppendLine(String.Format("Variable {0} is not defined.", identifier));

            if (output != null)
                output.Write(str.ToString());
            Write(str.ToString());

            if (output != null) output.Close();
        }

        private void DumpAttributes(IEnumerable<IPersistIfcEntity> entities, IEnumerable<string> attrNames, string outputPath = null, string identifier = null)
        {
            if (outputPath == null)
                ExportCSV(entities, attrNames, null);
            else
            {
                try
                {
                    var ext = Path.GetExtension(outputPath).ToLower();
                    if (ext == ".xls")
                        ExportXLS(entities, attrNames, outputPath, identifier);
                    else
                        ExportCSV(entities, attrNames, outputPath);
                    FileReportCreated(outputPath);
                }
                catch (Exception)
                {
                    
                    throw;
                }
            }
        }

        private void ExportCSV(IEnumerable<IPersistIfcEntity> entities, IEnumerable<string> attrNames, string outputPath = null)
        {
            TextWriter output = null;
            StringBuilder str = null;
            try
            {
                if (outputPath != null)
                {
                    output = new StreamWriter(outputPath, false);
                }

                str = new StringBuilder();
                var header = "";
                foreach (var name in attrNames)
                {
                    header += "," + name;
                }
                if (header.Length > 0)
                {
                    header = header.Remove(0, 1);
                    str.AppendLine(header);
                }

                foreach (var entity in entities)
                {
                    var line = "";
                    foreach (var name in attrNames)
                    {
                        //get attribute
                        var attr = GetAttributeValue(name, entity);
                        if (attr == null)
                            attr = GetPropertyValue(name, entity);
                        if (attr != null)
                            line += "," + attr.ToString();
                        else
                            line += ", - ";
                    }
                    if (line.Length > 0)
                    {
                        line = line.Remove(0, 1);
                        str.AppendLine(line);
                    }
                }

                if (output != null)
                    output.Write(str.ToString());
            }
            catch (Exception e)
            {
                Scanner.yyerror("It was not possible to dump/export specified content: " + e.Message);
            }
            finally
            {
                if (output != null)
                {
                    output.Close();
                }
                else
                    Write(str.ToString());
            }
        }

        private void ExportXLS(IEnumerable<IPersistIfcEntity> entities, IEnumerable<string> attrNames, string outputPath, string identifier)
        {
            if (String.IsNullOrEmpty(outputPath)) 
                throw new ArgumentException();

            if (identifier == null)
                identifier = "Exported_values";

            HSSFWorkbook workbook = null;
            try
            {
                if (File.Exists(outputPath))
                    workbook = new HSSFWorkbook(File.Open(outputPath, FileMode.Open, FileAccess.ReadWrite));
                else
                    workbook = new HSSFWorkbook();

                // Getting the worksheet by its name... 
                ISheet sheet = workbook.GetSheet(identifier) ?? workbook.CreateSheet(identifier);

                //create header
                var headerColour = GetColor(workbook, 220, 220, 220);
                var headerStyle = GetCellStyle(workbook, "", headerColour);
                var rowNum = sheet.LastRowNum == 0 ? 0 : sheet.LastRowNum + 1;
                IRow dataRow = sheet.CreateRow(rowNum);
                var cellType = dataRow.CreateCell(0, CellType.STRING);
                cellType.SetCellValue("IFC Type");
                cellType.CellStyle = headerStyle;
                var cellLabel = dataRow.CreateCell(1, CellType.STRING);
                cellLabel.SetCellValue("Label");
                cellLabel.CellStyle = headerStyle;

                var names = attrNames.ToList();
                for (int i = 0; i < attrNames.Count(); i++)
                {
                    var cell = dataRow.CreateCell(i + 2, CellType.STRING);
                    cell.SetCellValue(names[i]);
                    cell.CellStyle = headerStyle;
                }
                    

                //export values
                var entList = entities.ToList();
                for (int j = 0; j < entList.Count; j++)
                {
                    var entity = entList[j];
                    dataRow = sheet.CreateRow(++rowNum);
                    dataRow.CreateCell(0, CellType.STRING).SetCellValue(entity.GetType().Name);
                    var label = "#" + entity.EntityLabel;
                    dataRow.CreateCell(1, CellType.STRING).SetCellValue(label);

                    for (int i = 0; i < attrNames.Count(); i++)
                    {
                        var name = names[i];
                        var cell = dataRow.CreateCell(i+2, CellType.STRING);
                        //get attribute
                        var value = GetAttributeValue(name, entity);
                        if (value == null)
                            value = GetPropertyValue(name, entity);
                        if (value != null)
                            cell.SetCellValue(value.ToString());
                    }
                }

                var file = File.Create(outputPath);
                if (workbook != null)
                    workbook.Write(file);
                file.Close();
            }
            catch (Exception e)
            {
                Scanner.yyerror("It was not possible to dump/export specified content: " + e.Message);
            }
        }

        private HSSFColor GetColor (HSSFWorkbook workbook, byte red, byte green, byte blue)
        {
            HSSFPalette palette = workbook.GetCustomPalette();
            HSSFColor colour = palette.FindSimilarColor(red, green, blue);
            if (colour == null)
            {
                // First 64 are system colours
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                {
                    NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE = 64;
                }
                NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE++;
                colour = palette.AddColor(red, green, blue);
            }
            return colour;
        }

        private HSSFCellStyle GetCellStyle(HSSFWorkbook workbook, string formatString, HSSFColor colour)
        {
            HSSFCellStyle cellStyle;
            cellStyle = workbook.CreateCellStyle() as HSSFCellStyle;

            HSSFDataFormat dataFormat = workbook.CreateDataFormat() as HSSFDataFormat;
            cellStyle.DataFormat = dataFormat.GetFormat(formatString);

            cellStyle.FillForegroundColor = colour.GetIndex();
            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

            //cellStyle.BorderBottom = BorderStyle.THIN;
            //cellStyle.BorderLeft = BorderStyle.THIN;
            //cellStyle.BorderRight = BorderStyle.THIN;
            //cellStyle.BorderTop = BorderStyle.THIN;

            return cellStyle;
        }

        private void ClearIdentifier(string identifier)
        {
            if (Variables.IsDefined(identifier))
            {
                Variables.Clear(identifier);
            }
        }

        private int CountEntities(IEnumerable<IPersistIfcEntity> entities)
        {
            var result = entities.Count();
            WriteLine(result.ToString());
            return result;
        }

        private double SumEntities(string attrOrPropName, Tokens attrType, IEnumerable<IPersistIfcEntity> entities)
        {
            double sum = double.NaN;
            foreach (var entity in entities)
            {
                double value = GetDoubleValue(attrOrPropName, attrType, entity);
                if (!double.IsNaN(value))
                {
                    if (double.IsNaN(sum)) sum = 0;
                    sum += value;
                }
            }

            WriteLine(sum.ToString());
            return sum;
        }

        private double MinEntities(string attrOrPropName, Tokens attrType, IEnumerable<IPersistIfcEntity> entities)
        {
            double min = double.NaN;
            foreach (var entity in entities)
            {
                double value = GetDoubleValue(attrOrPropName, attrType, entity);
                if (!double.IsNaN(value))
                {
                    if (double.IsNaN(min)) 
                        min = value;
                    else
                        min = Math.Min(value, min);
                }
            }

            WriteLine(min.ToString());
            return min;
        }

        private double MaxEntities(string attrOrPropName, Tokens attrType, IEnumerable<IPersistIfcEntity> entities)
        {
            double max = double.NaN;
            foreach (var entity in entities)
            {
                double value = GetDoubleValue(attrOrPropName, attrType, entity);
                if (!double.IsNaN(value))
                {
                    if (double.IsNaN(max)) 
                        max = value;
                    else
                        max = Math.Max(value, max);
                }
            }

            WriteLine(max.ToString());
            return max;
        }

        private double AverageEntities(string attrOrPropName, Tokens attrType, IEnumerable<IPersistIfcEntity> entities)
        {
            double avg = double.NaN;
            foreach (var entity in entities)
            {
                double value = GetDoubleValue(attrOrPropName, attrType, entity);
                if (!double.IsNaN(value))
                {
                    if (double.IsNaN(avg)) avg = 0;
                    avg += value;
                }
            }

            if (!double.IsNaN(avg) && entities.Count() != 0)
            {
                avg = avg / entities.Count();
                WriteLine(avg.ToString());
                return avg;
            }

            return Double.NaN;
        }



        private double GetDoubleValue(string attrOrPropName, Tokens attrType, IPersistIfcEntity entity)
        {
            double value = double.NaN;
            switch (attrType)
            {
                case Tokens.STRING:
                    var attr1 = GetAttributeValue(attrOrPropName, entity);
                    if (attr1 == null)
                        attr1 = GetPropertyValue(attrOrPropName, entity);
                    if (attr1 != null)
                    {
                        var attrStr = attr1.ToString();
                        if (!Double.TryParse(attrStr, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                            value = Double.NaN;
                    }
                    break;
                case Tokens.ATTRIBUTE:
                    var attr = GetAttributeValue(attrOrPropName, entity);
                    if (attr != null)
                    {
                        var attrStr = attr.ToString();
                        if (!Double.TryParse(attrStr, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                            value = Double.NaN;
                    }
                    break;
                case Tokens.PROPERTY:
                    var prop = GetPropertyValue(attrOrPropName, entity);
                    if (prop != null)
                    {
                        var attrStr = prop.ToString();
                        if (!Double.TryParse(attrStr, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                            value = Double.NaN;
                    }
                    break;
                default:
                    throw new ArgumentException("Unexpected attribute type. STRING, ATTRIBUTE or PROPERTY expected");
            }
            return value;
        }

        #endregion

        #region Add or remove elements to and from group or type or spatial element
        private void AddOrRemove(Tokens action, IEnumerable<IPersistIfcEntity> entities, string aggregation)
        { 
        //conditions
            if (!Variables.IsDefined(aggregation))
            {
                Scanner.yyerror("Variable '" + aggregation + "' is not defined and doesn't contain any products.");
                return;
            }
            if (Variables[aggregation].Count() != 1)
            {
                Scanner.yyerror("Exactly one group, system, type or spatial element should be in '" + aggregation + "'.");
                return;
            }

            //check if all of the objects are from the actual model and not just referenced ones
            foreach (var item in entities)
            {
                if (item.ModelOf != _model)
                {
                    Scanner.yyerror("There is an object which is from referenced model so it cannot be used in this expression. Operation canceled.");
                    return;
                }
            }
            if (Variables[aggregation].FirstOrDefault().ModelOf != _model)
            {
                Scanner.yyerror("There is an object which is from referenced model so it cannot be used in this expression. Operation canceled.");
                return;
            }

            IfcGroup group = Variables[aggregation].FirstOrDefault() as IfcGroup;
            IfcTypeObject typeObject = Variables[aggregation].FirstOrDefault() as IfcTypeObject;
            IfcSpatialStructureElement spatialStructure = Variables[aggregation].FirstOrDefault() as IfcSpatialStructureElement;
            IfcClassificationNotationSelect classification = Variables[aggregation].FirstOrDefault() as IfcClassificationNotationSelect;

            if (group == null && typeObject == null && spatialStructure == null && classification == null)
            {
                Scanner.yyerror("Only 'group', 'system', 'spatial element' or 'type object' should be in '" + aggregation + "'.");
                return;
            }
            
            //Action which will be performed
            Action perform = null;

            if (classification != null)
            {
                var objects = entities.OfType<IfcRoot>().Cast<IfcRoot>();
                if (objects.Count() != entities.Count())
                    Scanner.yyerror("Only objects which are subtypes of 'IfcRoot' can be assigned to classification '" + aggregation + "'.");

                perform = () => {
                    foreach (var obj in objects)
                    {

                        switch (action)
                        {
                            case Tokens.ADD:
                                classification.AddObjectToClassificationNotation(obj);
                                break;
                            case Tokens.REMOVE:
                                classification.RemoveObjectFromClassificationNotation(obj);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException("Unexpected action. Only ADD or REMOVE can be used in this context.");
                        }
                    }
                };
            }

            if (group != null)
            {
                var objects = entities.OfType<IfcObjectDefinition>().Cast<IfcObjectDefinition>();
                if (objects.Count() != entities.Count())
                    Scanner.yyerror("Only objects which are subtypes of 'IfcObjectDefinition' can be assigned to group '" + aggregation + "'.");

                perform = () =>
                {
                    foreach (var obj in objects)
                    {

                        switch (action)
                        {
                            case Tokens.ADD:
                                group.AddObjectToGroup(obj);
                                break;
                            case Tokens.REMOVE:
                                group.RemoveObjectFromGroup(obj);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException("Unexpected action. Only ADD or REMOVE can be used in this context.");
                        }
                    }
                };
            }

            if (typeObject != null)
            {
                var objects = entities.OfType<IfcObject>().Cast<IfcObject>();
                if (objects.Count() != entities.Count())
                    Scanner.yyerror("Only objects which are subtypes of 'IfcObject' can be assigned to 'IfcTypeObject' '" + aggregation + "'.");

                perform = () => {
                    foreach (var obj in objects)
                    {
                        switch (action)
                        {
                            case Tokens.ADD:
                                obj.SetDefiningType(typeObject, _model);
                                //if there is material layer set defined for the type material layer set usage should be defined for the elements
                                var lSet = typeObject.GetMaterial() as IfcMaterialLayerSet;
                                if (lSet != null)
                                {
                                    var usage = _model.Instances.New<IfcMaterialLayerSetUsage>(u => {
                                        u.ForLayerSet = lSet;
                                        u.DirectionSense = IfcDirectionSenseEnum.POSITIVE;
                                        u.LayerSetDirection = IfcLayerSetDirectionEnum.AXIS1;
                                        u.OffsetFromReferenceLine = 0;
                                    });
                                    obj.SetMaterial(usage);
                                }
                                break;
                            case Tokens.REMOVE:
                                IfcRelDefinesByType rel = _model.Instances.Where<IfcRelDefinesByType>(r => r.RelatingType == typeObject && r.RelatedObjects.Contains(obj)).FirstOrDefault();
                                if (rel != null) rel.RelatedObjects.Remove(obj);
                                //remove material layer set usage if any exist. It is kind of indirect relation.
                                var lSet2 = typeObject.GetMaterial() as IfcMaterialLayerSet;
                                var usage2 = obj.GetMaterial() as IfcMaterialLayerSetUsage;
                                if (lSet2 != null && usage2 != null && usage2.ForLayerSet == lSet2)
                                {
                                    //the best would be to delete usage2 from the model but that is not supported bz the XbimModel at the moment
                                    var rel2 = _model.Instances.Where<IfcRelAssociatesMaterial>(r => 
                                            r.RelatingMaterial as IfcMaterialLayerSetUsage == usage2 && 
                                            r.RelatedObjects.Contains(obj)
                                        ).FirstOrDefault();
                                    if (rel2 != null) rel2.RelatedObjects.Remove(obj);
                                }
                                break;
                            default:
                                throw new ArgumentOutOfRangeException("Unexpected action. Only ADD or REMOVE can be used in this context.");
                        }
                    }
                };
            }

            if (spatialStructure != null)
            {
                var objects = entities.OfType<IfcProduct>().Cast<IfcProduct>();
                if (objects.Count() != entities.Count())
                    Scanner.yyerror("Only objects which are subtypes of 'IfcProduct' can be assigned to 'IfcSpatialStructureElement' '" + aggregation + "'.");

                perform = () =>
                {
                    foreach (var obj in objects)
                    {
                        switch (action)
                        {
                            case Tokens.ADD:
                                spatialStructure.AddElement(obj);
                                break;
                            case Tokens.REMOVE:
                                IfcRelContainedInSpatialStructure rel = _model.Instances.Where<IfcRelContainedInSpatialStructure>(r => r.RelatingStructure == spatialStructure && r.RelatedElements.Contains(obj)).FirstOrDefault();
                                if (rel != null) rel.RelatedElements.Remove(obj);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException("Unexpected action. Only ADD or REMOVE can be used in this context.");
                        }
                    }
                };
            }

            if (perform == null) return;

            //perform action
            if (!_model.IsTransacting)
            {
                using (var txn = _model.BeginTransaction("Group manipulation"))
                {
                    perform();
                    txn.Commit();
                }
            }
            else
                perform();
        }
        #endregion

        #region Model manipulation
        public void OpenModel(string path)
        {
            if (!File.Exists(path))
            {
                Scanner.yyerror("File doesn't exist: " + path);
                return;
            }
            try
            {
                string ext = Path.GetExtension(path).ToLower();
                if (ext == ".xbim" || ext == ".xbimf")
                    _model.Open(path, XbimExtensions.XbimDBAccess.Read,null);
                else
                   _model.CreateFrom(path, null, null, true,true);
                _model.CacheStart();
                ModelChanged(_model);
            }
            catch (Exception e)
            {
                Scanner.yyerror("File '"+path+"' can't be used as an input file. Model was not opened: " + e.Message);
            }
        }

        public void CloseModel()
        {
            try
            {
                _model.Close();
                _variables.Clear();
                _model = XbimModel.CreateTemporaryModel();
                ModelChanged(_model);
            }
            catch (Exception e)
            {

                Scanner.yyerror("Model could not have been closed: " + e.Message);
            }
            
        }

        public void ValidateModel()
        {
            try
            {
                TextWriter errOutput = new StringWriter();
                var errCount = _model.Validate(errOutput, XbimExtensions.ValidationFlags.Properties);

                if (errCount != 0)
                {
                    WriteLine("Number of errors: " + errCount.ToString());
                    WriteLine(errOutput.ToString());
                }
                else
                    WriteLine("No errors in the model.");
            }
            catch (Exception e)
            {
                Scanner.yyerror("Model could not be validated: " + e.Message);
            }
        }

        public void SaveModel(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
                _model.SaveAs(path);
            }
            catch (Exception e)
            {
                Scanner.yyerror("Model was not saved: " + e.Message);   
            }
        }

        public void AddReferenceModel(string refModel, string organization, string owner)
        {
            _model.AddModelReference(refModel, organization, owner);
            ModelChanged(_model);
        }

        public void CopyToModel(string variable, string model)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Objects manipulation
        private void EvaluateSetExpression(IEnumerable<IPersistIfcEntity> entities, IEnumerable<Expression> expressions)
        {
            Action perform = () => {
                if (entities == null) return;
                foreach (var expression in expressions)
                {
                    try
                    {
                        var action = Expression.Lambda<Action<IPersistIfcEntity>>(expression, _input).Compile();
                        entities.ToList().ForEach(action);
                    }
                    catch (Exception e)
                    {
                        Scanner.yyerror(e.Message);
                    }
                }    
            };

            if (_model.IsTransacting)
                perform();
            else
                using (var txn = _model.BeginTransaction("Setting properties and attribues"))
                {
                    perform();
                    txn.Commit();
                }
        }

        private Expression GenerateSetExpression(string attrName, object newVal, Tokens type)
        {
            var nameExpr = Expression.Constant(attrName);
            var valExpr = Expression.Convert(Expression.Constant(newVal), typeof(object));
            var thisExpr = Expression.Constant(this);

            string methodName = null;
            switch (type)
            {
                case Tokens.STRING:
                    methodName = "SetAttributeOrProperty";
                    break;
                case Tokens.PROPERTY:
                    methodName = "SetProperty";
                    break;
                case Tokens.ATTRIBUTE:
                    methodName = "SetAttribute";
                    break;
                default:
                    throw new ArgumentException("Unexpected value of the type. STRING, PROPERTY and ATTRIBUTE tokens expected only.");
            }

            var evaluateMethod = GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.NonPublic);
            return Expression.Call(thisExpr, evaluateMethod, _input, nameExpr, valExpr);
        }

        private void SetAttributeOrProperty(IPersistIfcEntity input, string attrName, object newVal)
        {
            //try to set attribute as a priority
            var attr = GetAttributeInfo(attrName, input);
            if (attr != null)
                SetValue(attr, input, newVal);
            else
                //set property if no such an attribute exist
                SetProperty(input, attrName, newVal);
        }

        private void SetAttribute(IPersistIfcEntity input, string attrName, object newVal)
        {
            if (input == null) return;

            var attr = GetAttributeInfo(attrName, input);
            SetValue(attr, input, newVal);
        }

        private void SetProperty(IPersistIfcEntity entity, string name, object newVal)
        {
            //try to get existing property
            object pObject = null;
            var pInfo = GetPropertyInfo(name, entity, out pObject);
            if (pInfo != null)
            {
            SetValue(pInfo, pObject, newVal);
            }

            //create new property if no such a property or quantity exists
            else
            {
                //try to get the name of the pSet if it is encoded in there
                var split = name.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                string pSetName = null;
                if (split.Count() == 2)
                {
                    pSetName = split[0];
                    name = split[1];
                }

                //prepare potential objects
                IfcObject obj = entity as IfcObject;
                IfcTypeObject typeObj = entity as IfcTypeObject;
                IfcMaterial material = entity as IfcMaterial;

                //set new property in specified or default property set
                pSetName = pSetName ?? Defaults.DefaultPSet;
                IfcValue val = null;
                if (newVal != null)
                    val = CreateIfcValueFromBasicValue(newVal, name);

                if (obj != null)
                {
                    obj.SetPropertySingleValue(pSetName, name, val);
                }
                else if (typeObj != null)
                {
                    typeObj.SetPropertySingleValue(pSetName, name, val);
                } 
                else if (material != null)
                {
                    material.SetExtendedSingleValue(pSetName, name, val);
                }
            }
        }

        private void SetValue(PropertyInfo info, object instance, object value)
        {
            if (info == null)
            {
                Scanner.yyerror("It is not possible to set value of the property or attribute which doesn't exist.");
                return;
            }
            try
            {
                if (value != null)
                {
                    var targetType = info.PropertyType.IsNullableType()
                        ? Nullable.GetUnderlyingType(info.PropertyType)
                        : info.PropertyType;

                    object newValue = null;
                    if (!targetType.IsInterface && !targetType.IsAbstract && !targetType.IsEnum)
                        newValue = Activator.CreateInstance(targetType, value);
                    else if (targetType.IsEnum)
                    {
                        //this can throw exception if the name is not correct
                        newValue = Enum.Parse(targetType, value.ToString(), true);
                    }
                    else
                        newValue = CreateIfcValueFromBasicValue(value, info.Name);
                    //this will throw exception if the types are not compatible
                    info.SetValue(instance, newValue, null);
                }
                else
                    //this can throw error if the property can't be null (like structure)
                    info.SetValue(instance, null, null);
            }
            catch (Exception)
            {
                throw new Exception("Value "+ (value != null ? value.ToString() : "NULL") +" could not have been set to "+ info.Name + " of type"+ instance.GetType().Name + ". Type should be compatible with " + info.MemberType);
            }
            
        }

        private static IfcValue CreateIfcValueFromBasicValue(object value, string propName)
        {
            Type type = value.GetType();
            if (type == typeof(int))
                return new IfcInteger((int)value);
            if (type == typeof(string))
                return new IfcText((string)value);
            if (type == typeof(double))
                return new IfcNumericMeasure((double)value);
            if (type == typeof(bool))
                return new IfcBoolean((bool)value);

            throw new Exception("Unexpected type of the new value " + type.Name + " for property " + propName);
        }

        private Expression GenerateSetMaterialExpression(string materialIdentifier)
        {
            var entities = _variables.GetEntities(materialIdentifier);

            if (entities == null)
            {
                Scanner.yyerror("There should be exactly one material in the variable " + materialIdentifier);
                return Expression.Empty();
            }
            var count = entities.Count();
            if (count != 1)
            {
                Scanner.yyerror("There should be only one object in the variable " + materialIdentifier);
                return Expression.Empty();
            }
            var material = entities.FirstOrDefault() as IfcMaterialSelect;
            if (material == null)
            {
                Scanner.yyerror("There should be exactly one material in the variable " + materialIdentifier);
                return Expression.Empty();
            }

            var materialExpr = Expression.Constant(material);
            var scanExpr = Expression.Constant(Scanner);

            var evaluateMethod = GetType().GetMethod("SetMaterial", BindingFlags.NonPublic|BindingFlags.Static);
            return Expression.Call(null, evaluateMethod, _input, materialExpr, scanExpr);
        }

        private static void SetMaterial(IPersistIfcEntity entity, IfcMaterialSelect material, AbstractScanner<ValueType, LexLocation> scanner)
        {
            if (entity == null || material == null) return;

            var materialSelect = material as IfcMaterialSelect;
            if (materialSelect == null)
            {
                scanner.yyerror(material.GetType() + " can't be used as a material");
                return;
            }
            var root = entity as IfcRoot;
            if (root == null)
            {
                scanner.yyerror(root.GetType() + " can't have a material assigned.");
                return;
            }


            IModel model = material.ModelOf;
            var matSet = material as IfcMaterialLayerSet;
            if (matSet != null)
            {
                var element = root as IfcElement;
                if (element != null)
                {
                    var usage = model.Instances.New<IfcMaterialLayerSetUsage>(mlsu => {
                        mlsu.DirectionSense = IfcDirectionSenseEnum.POSITIVE;
                        mlsu.ForLayerSet = matSet;
                        mlsu.LayerSetDirection = IfcLayerSetDirectionEnum.AXIS1;
                        mlsu.OffsetFromReferenceLine = 0;
                    });
                    var rel = model.Instances.New<IfcRelAssociatesMaterial>(r => {
                        r.RelatedObjects.Add(root);
                        r.RelatingMaterial = usage;
                    });
                    return;
                }
            }

            var matUsage = material as IfcMaterialLayerSetUsage;
            if (matUsage != null)
            {
                var typeElement = root as IfcElementType;
                if (typeElement != null)
                {
                    //change scope to the layer set for the element type. It will be processed in a standard way than
                    materialSelect = matUsage.ForLayerSet;
                }
            }

            //find existing relation
            var matRel = model.Instances.Where<IfcRelAssociatesMaterial>(r => r.RelatingMaterial == materialSelect).FirstOrDefault();
            if (matRel == null)
                //create new if none exists
                matRel = model.Instances.New<IfcRelAssociatesMaterial>(r => r.RelatingMaterial = materialSelect);
            //insert only if it is not already there
            if (!matRel.RelatedObjects.Contains(root)) matRel.RelatedObjects.Add(root);

        }
        #endregion

        #region Thickness conditions
        private Expression GenerateThicknessCondition(double thickness, Tokens condition)
        {
            var thickExpr = Expression.Constant(thickness);
            var condExpr = Expression.Constant(condition);

            var evaluateMethod = GetType().GetMethod("EvaluateThicknessCondition", BindingFlags.Static | BindingFlags.NonPublic);
            return Expression.Call(null, evaluateMethod, _input, thickExpr, condExpr);
        }

        private static bool EvaluateThicknessCondition(IPersistIfcEntity input, double thickness, Tokens condition)
        {
            IfcRoot root = input as IfcRoot;
            if (input == null) return false;

            double? value = null;
            var materSel = root.GetMaterial();
            IfcMaterialLayerSetUsage usage = materSel as IfcMaterialLayerSetUsage;
            if (usage != null)
                if (usage.ForLayerSet != null) 
                    value = usage.ForLayerSet.MaterialLayers.Aggregate(0.0, (current, layer) => current + layer.LayerThickness);
            IfcMaterialLayerSet set = materSel as IfcMaterialLayerSet;
            if (set != null)
                value = set.TotalThickness;
            if (value == null)
                return false;
            switch (condition)
            {
                case Tokens.OP_EQ:
                    return thickness.AlmostEquals(value ?? 0);
                case Tokens.OP_NEQ:
                    return !thickness.AlmostEquals(value ?? 0);
                case Tokens.OP_GT:
                    return value > thickness;
                case Tokens.OP_LT:
                    return value < thickness;
                case Tokens.OP_GTE:
                    return value >= thickness;
                case Tokens.OP_LTQ:
                    return value <= thickness;
                default:
                    throw new ArgumentException("Unexpected value of the condition");
            }
        }
        #endregion

        #region Summary

        public void ExportSummary(string path, string format="json")
        {
            if (string.Compare(format, "json", true) == 0)
            {
                string outputFile = Path.ChangeExtension(path, ".json"); //enforce JSON
                using (StreamWriter sw = new StreamWriter(path))
                {
                    using (JsonTextWriter jw = new JsonTextWriter(sw))
                    {
                        JsonSerializer s = JsonSerializer.Create();
                        s.Serialize(jw, new XbimModelSummary(_model));
                    }
                }
            }
        }

        #endregion


        #region COBie
        public void ExportCOBie(string path, string template)
        {
            string outputFile = Path.ChangeExtension(path, ".xls"); //enforce xls
            FilterValues UserFilters = new FilterValues();//COBie Class filters, set to initial defaults
            // Build context
            COBieContext context = new COBieContext();
            context.TemplateFileName = template;
            context.Model = _model;
            //set filter option
            context.Exclude = UserFilters;

            //set the UI language to get correct resource file for template
            //if (Path.GetFileName(parameters.TemplateFile).Contains("-UK-"))
            //{
            try
            {
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-GB");
                System.Threading.Thread.CurrentThread.CurrentUICulture = ci;
            }
            catch (Exception)
            {
                //to nothing Default culture will still be used

            }

            COBieBuilder builder = new COBieBuilder(context);
            COBieXLSSerialiser serialiser = new COBieXLSSerialiser(outputFile, context.TemplateFileName);
        
            serialiser.Excludes = UserFilters;
            builder.Export(serialiser);
        }


        #endregion

        #region Creation of classification systems
        private void CreateClassification(string name)
        {
            ClassificationCreator creator = new ClassificationCreator();

            try
            {
                creator.CreateSystem(_model, name);
            }
            catch (Exception e)
            {
                Scanner.yyerror("Classification {0} couldn't have been created: {1}", name, e.Message);
            }
        }
        #endregion

        #region Group conditions
        private Expression GenerateGroupCondition(Expression expression) 
        {
            var function = Expression.Lambda<Func<IPersistIfcEntity, bool>>(expression, _input).Compile();
            var fceExpr = Expression.Constant(function);

            var evaluateMethod = GetType().GetMethod("EvaluateGroupCondition", BindingFlags.Static | BindingFlags.NonPublic);

            return Expression.Call(null, evaluateMethod, _input, fceExpr);
        }

        private static bool EvaluateGroupCondition(IPersistIfcEntity input, Func<IPersistIfcEntity, bool> function)
        {
            foreach (var item in GetGroups(input))
            {
                if (function(item)) return true;
            }
            return false;
        }

        private static IEnumerable<IfcGroup> GetGroups(IPersistIfcEntity input)
        {
            IModel model = input.ModelOf;
            var obj = input as IfcObjectDefinition;
            if (obj != null)
            {
                var rels = model.Instances.Where<IfcRelAssignsToGroup>(r => r.RelatedObjects.Contains(input));
                foreach (var rel in rels)
                {
                    yield return rel.RelatingGroup;

                    //recursive search for upper groups in the hierarchy
                    foreach (var gr in GetGroups(rel.RelatingGroup))
                    {
                        yield return gr;
                    }
                }
            }
        }
        #endregion

        #region Spatial conditions
        private Expression GenerateSpatialCondition(Tokens op, Tokens condition, string identifier)
        {
            IEnumerable<IfcProduct> right = _variables[identifier].OfType<IfcProduct>();
            if (right.Count() == 0)
            {
                Scanner.yyerror("There are no suitable objects for spatial condition in " + identifier + ".");
                return Expression.Empty();
            }
            if (right.Count() != _variables[identifier].Count())
                Scanner.yyerror("Some of the objects in " + identifier + " can't be in a spatial condition.");

            var rightExpr = Expression.Constant(right);
            var condExpr = Expression.Constant(condition);
            var opExpr = Expression.Constant(op);
            var scanExpr = Expression.Constant(Scanner);
            var thisExpr = Expression.Constant(this);

            var evaluateMethod = GetType().GetMethod("EvaluateSpatialCondition", BindingFlags.NonPublic | BindingFlags.Instance);

            return Expression.Call(thisExpr, evaluateMethod, _input, opExpr, condExpr, rightExpr, scanExpr);
        }

        private bool EvaluateSpatialCondition(IPersistIfcEntity input, Tokens op, Tokens condition, IEnumerable<IfcProduct> right)
        {
            IfcProduct left = input as IfcProduct;
            if (left == null)
            {
                Scanner.yyerror(input.GetType().Name + " can't have a spatial condition.");
                return false;
            }

            switch (condition)
            {
                case Tokens.NORTH_OF:
                    break;
                case Tokens.SOUTH_OF:
                    break;
                case Tokens.WEST_OF:
                    break;
                case Tokens.EAST_OF:
                    break;
                case Tokens.ABOVE:
                    break;
                case Tokens.BELOW:
                    break;
                case Tokens.SPATIALLY_EQUALS:
                    break;
                case Tokens.DISJOINT:
                    break;
                case Tokens.INTERSECTS:
                    break;
                case Tokens.TOUCHES:
                    break;
                case Tokens.CROSSES:
                    break;
                case Tokens.WITHIN:
                    break;
                case Tokens.SPATIALLY_CONTAINS:
                    break;
                case Tokens.OVERLAPS:
                    break;
                case Tokens.RELATE:
                    break;
                default:
                    break;
            }

            throw new NotImplementedException();
        }
        #endregion


        #region Existance conditions
        private Expression GenerateExistanceCondition(Tokens existanceCondition, string modelName)
        {
            throw new NotImplementedException();
        }

        private Expression GenerateExistanceCondition(Tokens conditionA, string modelA, Tokens conditionB, string modelB)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Rule checking
        private RuleCheckResultsManager _ruleChecks = new RuleCheckResultsManager();
        public RuleCheckResultsManager RuleChecks { get { return _ruleChecks; } }

        private void CheckRule(string ruleName, Expression condition, IEnumerable<IPersistIfcEntity> elements)
        {
            var func = Expression.Lambda<Func<IPersistIfcEntity, bool>>(condition, _input).Compile();
            foreach (var item in elements)
            {
                //check the rule
                var check = func(item);
                var result = new RuleCheckResult() { Entity = item, IsCompliant = check};
                _ruleChecks.Add(ruleName, result);

                if (!check)
                    WriteLine(String.Format("{0} #{1} doesn't comply to the rule \"{2}\"", item.GetType().Name, item.EntityLabel, ruleName));
            }
        }

        private void CheckRule(string ruleName, double valueA, Tokens condition, double valueB)
        {
                //check the rule
                var check = CheckDoubleValue(valueA, valueB, condition);
                var result = new RuleCheckResult() { Entity = null, IsCompliant = check };
                _ruleChecks.Add(ruleName, result);

                if (!check)
                    WriteLine(String.Format("Values don't comply to the rule \"{0}\"", ruleName));
        }

        private bool CheckDoubleValue(double valueA, double valueB, Tokens condition)
        {
            if (Double.IsNaN(valueA) || Double.IsNaN(valueB))
                return false;

            switch (condition)
            {
                case Tokens.OP_GT:
                    return valueA > valueB;
                case Tokens.OP_LT:
                    return valueA < valueB;
                case Tokens.OP_GTE:
                    return valueA >= valueB;
                case Tokens.OP_LTQ:
                    return valueA <= valueB;
                case Tokens.OP_EQ:
                    return Math.Abs( valueA - valueB) < 1E-9;
                case Tokens.OP_NEQ:
                    return Math.Abs(valueA - valueB) > 1E-9;
                default:
                    throw new NotImplementedException("Unexpected token value");
            }
        
        }

        private void ClearRules()
        {
            _ruleChecks = new RuleCheckResultsManager();
        }

        private void SaveRules(string path)
        {
            var ext = (Path.GetExtension(path) ?? "").ToLower();
            try
            {
                string finalPath = null;
                if (ext == ".xls")
                    finalPath = _ruleChecks.SaveToXLS(path);
                else
                    finalPath = _ruleChecks.SaveToCSV(path);
                FileReportCreated(finalPath);
            }
            catch (Exception e)
            {
                Scanner.yyerror("It was not possible to save rule checking results to the file {0}. Error: {1}", path, e.Message);
            }

        }

        #endregion

        /// <summary>
        /// Unified function so that the same output can be send 
        /// to the Console and to the optional text writer as well.
        /// </summary>
        /// <param name="message">Message for the output</param>
        private void WriteLine(string message)
        {
            Console.WriteLine(message);
            if (Output != null)
                Output.WriteLine(message);
        }

        /// <summary>
        /// Unified function so that the same output can be send 
        /// to the Console and to the optional text writer as well.
        /// </summary>
        /// <param name="message">Message for the output</param>
        private void Write(string message)
        {
            Console.Write(message);
            if (Output != null)
                Output.Write(message);
        }


        #region Events
        /// <summary>
        /// This event is fired when parser open or 
        /// close model from script action
        /// </summary>
        public event ModelChangedHandler OnModelChanged;
        private void ModelChanged(XbimModel newModel)
        {
            if (OnModelChanged != null)
                OnModelChanged(this, new ModelChangedEventArgs(newModel));
        }

        /// <summary>
        /// This event is fired when file with report is created 
        /// (like XLS or CSV as an export of properties and arguments 
        /// or result of rule checking)
        /// </summary>
        public event FileReportCreatedHandler OnFileReportCreated;
        private void FileReportCreated(string path)
        {
            if (OnFileReportCreated != null)
                OnFileReportCreated(this, new FileReportCreatedEventArgs(path));
        }
        #endregion
    }

    internal struct Layer
    {
        public string material;
        public double thickness;
    }

    public static class TypeExtensions
    {
        public static bool IsNullableType(this Type type)
        {
            return type.IsGenericType
            && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }

        public static bool AlmostEquals(this double number, double value)
        {
            return Math.Abs(number - value) < 0.000000001;
        }
    }

    public delegate void ModelChangedHandler(object sender, ModelChangedEventArgs e);

    public class ModelChangedEventArgs : EventArgs
    {
        private XbimModel _newModel;
        public XbimModel NewModel { get { return _newModel; } }
        public ModelChangedEventArgs(XbimModel newModel)
        {
            _newModel = newModel;
        }
    }

    public delegate void FileReportCreatedHandler(object sender, FileReportCreatedEventArgs e);

    public class FileReportCreatedEventArgs : EventArgs
    {
        private string _filePath;
        public string FilePath { get { return _filePath; } }
        public FileReportCreatedEventArgs(string path)
        {
            _filePath = path;
        }
    }

    #region Rule checking results infrastructure
    public class RuleCheckResult
    {
        public IPersistIfcEntity Entity { get; set; }
        public bool IsCompliant { get; set; }
    }

    public class RuleCheckResults
    {
        public string Rule { get; set; }
        public List<RuleCheckResult> CheckResults { get; set; }

        public RuleCheckResults()
        {
            CheckResults = new List<RuleCheckResult>();
        }


    }

    public class RuleCheckResultsManager
    {
        private List<RuleCheckResults> _results = new List<RuleCheckResults>();
        public IEnumerable<RuleCheckResults> Results
        {
            get 
            {
                foreach (var item in _results)
                {
                    yield return item;
                }
            }
        }

        public void Add(string rule, RuleCheckResult result)
        {
            var res = _results.Where(r => r.Rule == rule);
            if (res.Any())
                res.FirstOrDefault().CheckResults.Add(result);
            else
            {
                var r = new RuleCheckResults() { Rule = rule };
                r.CheckResults.Add(result);
                _results.Add(r);
            }
        }

        public void Clear()
        {
            _results = new List<RuleCheckResults>();
        }

        /// <summary>
        /// Saves report of the rules check as a CSV file
        /// </summary>
        /// <param name="path">Path of the file</param>
        /// <returns>Final path with extension</returns>
        public string SaveToCSV(string path)
        {
            try
            {
                var ext = Path.GetExtension(path);
                if (String.IsNullOrEmpty(ext))
                    path += ".csv";

                //create text file or overwrite existing
                var file = File.CreateText(path);

                //write header
                file.WriteLine("{0},{1},{2},{3}", "Rule", "Element type", "Label", "Is Compliant");

                //write results
                foreach (var result in _results)
                    foreach (var ruleRes in result.CheckResults)
                    {
                        var entity = ruleRes.Entity;
                        var label = entity == null ? 0 : entity.EntityLabel;
                        var typeName = entity == null ? "" : entity.GetType().Name;
                        file.WriteLine("{0},{1},#{2},{3}", result.Rule, typeName, label, ruleRes.IsCompliant);
                    }

                //close file
                file.Close();
                return path;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Saves report of the rules check as a XLS file
        /// </summary>
        /// <param name="path">Path of the file</param>
        /// <returns>Final path with extension</returns>
        public string SaveToXLS(string path)
        {
            try
            {
                var ext = Path.GetExtension(path);
                if (String.IsNullOrEmpty(ext))
                    path += ".xls";

                //create text file or overwrite existing
                var rowNum = -1;
                var cellNum = -1;
                HSSFWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Rule_checking_results");
                IRow dataRow = sheet.CreateRow(++rowNum);
                

                //write header
                var cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                cell.SetCellValue("Rule");
                cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                cell.SetCellValue("Element type");
                cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                cell.SetCellValue("Label");
                cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                cell.SetCellValue("Is Compliant");

                //write results
                foreach (var result in _results)
                    foreach (var ruleRes in result.CheckResults)
                    {
                        var entity = ruleRes.Entity;
                        var label = entity == null ? 0 : entity.EntityLabel;
                        var typeName = entity == null ? "" : entity.GetType().Name;
                        dataRow = sheet.CreateRow(++rowNum);
                        cellNum = -1;

                        cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                        cell.SetCellValue(result.Rule);
                        cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                        cell.SetCellValue(typeName);
                        cell = dataRow.CreateCell(++cellNum, CellType.STRING);
                        cell.SetCellValue("#" + label);
                        cell = dataRow.CreateCell(++cellNum, CellType.BOOLEAN);
                        cell.SetCellValue(ruleRes.IsCompliant);
                    }

                //save and close file
                var file = File.Create(path);
                if (workbook != null)
                    workbook.Write(file);
                file.Close();
                return path;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
    #endregion
}
