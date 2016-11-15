using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xbim.XbimExtensions.Interfaces;
using System.Text.RegularExpressions;
using Xbim.Common;

namespace Xbim.Script
{
    public class XbimVariables
    {

        private Dictionary<string, IEnumerable<IPersistEntity>> _data = new Dictionary<string, IEnumerable<IPersistEntity>>();
        private string _lastVariable = null;
        public string LastVariable { get { return _lastVariable; } }
        public IEnumerable<IPersistEntity> LastEntities
        {
            get 
            {
                if (_lastVariable == null || !IsDefined(_lastVariable)) return new IPersistEntity[] { };
                return this[_lastVariable];
            }
        }
        
        public IEnumerable<IPersistEntity> GetEntities(string variable)
        {
            FixVariableName(ref variable);
            IEnumerable<IPersistEntity> result = new IPersistEntity[] { };
            _data.TryGetValue(variable, out result);
            return result;
        }

        public IEnumerable<string> GetVariables
        {
            get
            {
                return _data.Select(x => x.Key).AsEnumerable();
            }
        }



        public bool IsDefined(string variable)
        {
            FixVariableName(ref variable);
            return _data.ContainsKey(variable);
        }

        public void Set(string variable, IPersistEntity entity)
        {
            FixVariableName(ref variable);
            if (entity == null && IsDefined(variable))
                Clear(variable);
            else
                Set(variable, new IPersistEntity[] { entity });
        }

        public void Set(string variable, IEnumerable<IPersistEntity> entities)
        {
            FixVariableName(ref variable);
            if (IsDefined(variable))
                _data[variable] = entities.ToList();
            else
                _data.Add(variable, entities.ToList());
            _lastVariable = variable;
        }

        public void AddEntities(string variable, IEnumerable<IPersistEntity> entities)
        {
            FixVariableName(ref variable);
            if (IsDefined(variable))
            {
                _data[variable] = _data[variable].Union(entities.ToList());
            }
            else
                _data.Add(variable, entities.ToList());

            _lastVariable = variable;
        }

        public void RemoveEntities(string variable, IEnumerable<IPersistEntity> entities)
        {
            FixVariableName(ref variable);
            if (IsDefined(variable))
            {
                _data[variable] = _data[variable].Except(entities.ToList());
            }
            else
                throw new ArgumentException("Can't remove entities from variable which is not defined.");

            _lastVariable = variable;
        }

        public IEnumerable<IPersistEntity> this[string key]
        {
            get
            {
                FixVariableName(ref key);
                if (_data.ContainsKey(key))
                    return _data[key];
                else
                    return new IPersistEntity [] {};
            }
        }

        public void Clear() 
        {
            _data.Clear();
        }

        public void Clear(string identifier)
        {
            FixVariableName(ref identifier);
            if (IsDefined(identifier))
                _data[identifier] = new IPersistEntity[] { };
            else
                throw new ArgumentException(identifier + " is not defined;");
        }

        private void FixVariableName(ref string variable)
        {
            if (String.IsNullOrEmpty(variable))
                throw new ArgumentNullException("variable");
            if (variable[0] != '$')
                variable = "$" + variable;
        }
    }
}
