using System;
using System.Collections.Generic;
using System.Linq;
using Xbim.Ifc2x3.Kernel;
using Xbim.Ifc2x3.Extensions;
using Xbim.Ifc2x3.MeasureResource;
using Xbim.IO;

namespace Xbim.Script
{
    public static class ScriptGroupExtension
    {
        private static string _pScript = "script";

        /// <summary>
        /// Sets the script to the predefined pSet. This can be used for a late evaluation of the group members
        /// </summary>
        /// <param name="script">Script to be executed. Set this to NULL to reset the script. 
        /// Property will still be defined but will be empty. You can use "ExecuteScript" 
        /// to get the entities which belong to this group. These entities are not defined by 
        /// IfcRelAssignToGroup relationship.</param>
        public static void SetScript(this IfcGroup group, string script)
        {
            if (String.IsNullOrEmpty(script))
                group.SetPropertySingleValue(Defaults.DefaultPSet, _pScript, typeof(IfcText));
            else
            {
                //check if the script contains '$group' to be used to get elements
                if (!script.Contains("$group"))
                    throw new ArgumentException("Script doesn't contain '$group' variable.");
                group.SetPropertySingleValue(Defaults.DefaultPSet, _pScript, new IfcText(script));
            }
        }

        /// <summary>
        /// Gets the script from predefined pSet and property.
        /// </summary>
        /// <returns>Script saved in the property</returns>
        public static string GetScript(this IfcGroup group)
        {
            return group.GetPropertySingleValue<IfcText>(Defaults.DefaultPSet, _pScript);
        }

        /// <summary>
        /// Indicates whether the script is defined for this group
        /// </summary>
        /// <returns>True if the script is defined, false otherwise.</returns>
        public static bool HasScript(this IfcGroup group)
        {
            string script = group.GetPropertySingleValue<IfcText>(Defaults.DefaultPSet, _pScript);
            return  !String.IsNullOrEmpty(script);
        }

        /// <summary>
        /// This function will execute the script if there is any defined. 
        /// If no script is defined this will return empty set.
        /// </summary>
        /// <returns>Set of the results or empty set if there are no results or no script defined.</returns>
        public static IEnumerable<IfcObjectDefinition> ExecuteScript(this IfcGroup group)
        {
            var script = group.GetScript();
            if (script == null)
            {
                _lastErrors = new List<string>();
                return new IfcObjectDefinition[] { };
            }
            var model = group.ModelOf as XbimModel;
            var parser = new XbimQueryParser(model);
            parser.Parse(script);
            _lastErrors = parser.Errors;

            return parser.Results.GetEntities("$group").OfType<IfcObjectDefinition>();
        }

        /// <summary>
        /// Gets last errors from execution of the script. 
        /// You should always check this ater you try to get elements using 'ExecuteScript()'
        /// </summary>
        /// <returns>List of last errors with descriptive informations</returns>
        public static List<string> GetLastScriptErrors()
        {
            return _lastErrors;
        }

        private static List<string> _lastErrors;
    }
}
