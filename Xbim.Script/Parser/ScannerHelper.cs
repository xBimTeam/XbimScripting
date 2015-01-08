using System;
using System.Collections.Generic;
using System.Linq;
using Xbim.Ifc2x3.Kernel;
using System.Globalization;
using Xbim.Ifc2x3.ActorResource;

namespace Xbim.Script
{
    internal sealed partial class Scanner
    {
        /// <summary>
        /// String processing funtion used during the scanning.
        /// </summary>
        /// <returns>Token PRODUCT, PRODUCT_TYPE or STRING</returns>
        private Tokens ProcessString()
        {
            yylval.strVal = yytext;

            SetValue();
            Type t = GetProductType(yytext);
            if (t != null)
            {
                yylval.typeVal = t;
                return Tokens.PRODUCT;
            }
            t = GetTypeProductType(yytext);
            if (t != null)
            {
                yylval.typeVal = t;
                return Tokens.PRODUCT_TYPE;
            }

            //return string otherwise
            return Tokens.STRING;
        }

        /// <summary>
        /// function used by scanner to set values for value type tokens
        /// </summary>
        /// <param name="type">Value type. If no value type is specified 'STRING' is used by default</param>
        /// <returns>Token set by the function</returns>
        private Tokens SetValue(Tokens type = Tokens.STRING)
        {
            yylval.strVal = yytext;

            switch (type)
            {
                case Tokens.INTEGER:
                    if (int.TryParse(yytext, out yylval.intVal))
                        return type;
                    break;
                case Tokens.DOUBLE:
                    try
                    {
                        yylval.doubleVal = double.Parse(yytext, CultureInfo.InvariantCulture);
                        return type;
                    }
                    catch (Exception)
                    {
                        
                    }
                    break;
                case Tokens.BOOLEAN:
                    if (yytext.ToLower() == ".t.")
                    {
                        yylval.boolVal = true;
                        return type;
                    }
                    if (yytext.ToLower() == ".f.")
                    {
                        yylval.boolVal = false;
                        return type;
                    }
                    if (bool.TryParse(yytext, out yylval.boolVal))
                        return type;
                    break;
                case Tokens.PRODUCT:
                    yylval.typeVal = null;
                    yylval.typeVal = GetProductType(yytext);
                    if (yylval.typeVal != null)
                        return type;
                    break;
                case Tokens.PRODUCT_TYPE:
                    yylval.typeVal = null;
                    yylval.typeVal = GetTypeProductType(yytext);
                    if (yylval.typeVal != null)
                        return type;
                    break;
                case Tokens.MATERIAL:
                    yylval.typeVal = typeof(Xbim.Ifc2x3.MaterialResource.IfcMaterial);
                    return type;
                case Tokens.GROUP:
                    yylval.typeVal = typeof(IfcGroup);
                    return type;
                case Tokens.ORGANIZATION:
                    yylval.typeVal = typeof(IfcOrganization);
                    return type;
                case Tokens.IDENTIFIER:
                    return type;
                default:
                    yylval.strVal = yytext.Trim('\'', '"');
                    return Tokens.STRING;
            }
            return Tokens.STRING;
        }


        /// <summary>
        /// Dictionary of possible names and types
        /// </summary>
        private static Dictionary<string, Type> ProductDictionary = GetNames(typeof(IfcObject));

        /// <summary>
        /// Dictionary of possible type product names and types
        /// </summary>
        private static Dictionary<string, Type> TypeProductDictionary = GetNames(typeof(IfcTypeProduct));

        /// <summary>
        /// Find if the name is one of the variants of some product name. Case insensitive.
        /// </summary>
        /// <param name="productName">Name of the product</param>
        /// <returns>True it the string can be name of some product, false otherwise</returns>
        private bool IsProduct(string productName)
        {
            return ProductDictionary.ContainsKey(productName.ToLower());
        }

        /// <summary>
        /// Find if the name is one of the variants of some type product name. Case insensitive.
        /// </summary>
        /// <param name="productName">Name of the type product</param>
        /// <returns>True it the string can be name of some type product, false otherwise</returns>
        private bool IsTypeProduct(string typeProductName)
        {
            return TypeProductDictionary.ContainsKey(typeProductName.ToLower());
        }

        /// <summary>
        /// Gets product type if the name is one of variants (case insensitive)
        /// </summary>
        /// <param name="productName">Name of the product (IfcWall, wall, IfcStairFlight, stair flight, ...)</param>
        /// <returns>Type if found, null otherwise</returns>
        private Type GetProductType(string productName)
        {
            Type result = null;
            ProductDictionary.TryGetValue(productName.ToLower(), out result);
            return result;
        }

        /// <summary>
        /// Gets type product type if the name is one of variants (case insensitive)
        /// </summary>
        /// <param name="typeProductName">Name of the type product (IfcWallType, wall_type, ...)</param>
        /// <returns>Type if found, null otherwise</returns>
        private Type GetTypeProductType(string typeProductName)
        {
            Type result = null;
            TypeProductDictionary.TryGetValue(typeProductName.ToLower(), out result);
            return result;
        }

        /// <summary>
        /// Creates dictionary where keys are variants of product names and values are the types
        /// </summary>
        /// <returns>Dictionary of variants of the product names</returns>
        private static Dictionary<string, Type> GetNames(Type baseType)
        {
            var assembly = baseType.Assembly;
            var types = assembly.GetTypes().Where(t => baseType.IsAssignableFrom(t));
            var result = new Dictionary<string, Type>();
            foreach (var type in types)
            {
                string name = type.Name;
                //plain IFC name
                result.Add(name.ToLower(), type);

                //IFC name without Ifc prefix
                string shortName = name.Remove(0, 3);
                result.Add(shortName.ToLower(), type);

                //IFC name without Ifc prefix and splitted according to camel case.
                string splitName = SplitCamelCase(name.Remove(0, 3));
                if (splitName != shortName)
                    result.Add(splitName.ToLower(), type);
            }
            return result;
        }

        /// <summary>
        /// Splits the string from camel case to underscore separated strings
        /// </summary>
        /// <param name="input">Input string</param>
        /// <returns>Splitted camel-case string</returns>
        private static string SplitCamelCase(string input)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, "([A-Z])", "_$1", System.Text.RegularExpressions.RegexOptions.Compiled).Trim('_');
        }

        /// <summary>
        /// List of errors
        /// </summary>
        public List<string> Errors = new List<string>();

        /// <summary>
        /// List of error locations
        /// </summary>
        public List<ErrorLocation> ErrorLocations = new List<ErrorLocation>();

        /// <summary>
        /// Overriden yyerror function for error reporting
        /// </summary>
        /// <param name="format">Formated error message</param>
        /// <param name="args">Error arguments</param>
        public override void yyerror(string format, params object[] args)
        {
            Errors.Add(String.Format(format, args) + String.Format("From line {0}, column {1} to line {2}, column {3}", yylloc.StartLine, yylloc.StartColumn, yylloc.EndLine, yylloc.EndColumn));
            ErrorLocations.Add(new ErrorLocation() { 
                StartLine = yylloc.StartLine, 
                EndLine = yylloc.EndLine, 
                StartColumn = yylloc.StartColumn, 
                EndColumn = yylloc.EndColumn,
                Message = String.Format(format, args)
            });

            base.yyerror(format, args);
        }


    }

    public struct ErrorLocation
	{
		public int StartLine;
        public int EndLine;
        public int StartColumn;
        public int EndColumn;
        public string Message;
	}
}
