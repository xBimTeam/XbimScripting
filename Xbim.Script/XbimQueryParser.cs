using System;
using System.Collections.Generic;
using Xbim.IO;

namespace Xbim.Script
{
    public class XbimQueryParser
    {
        private Parser _parser;
        private Scanner _scanner;

        private string _strSource;
        private IList<string> _strListSource;
        private System.IO.Stream _streamSource;

        /// <summary>
        /// Errors encountered during parsing.
        /// </summary>
        public List<string> Errors { get { return _scanner.Errors; } }

        /// <summary>
        /// Locations of the errors. ErrorLocation contains error message and
        /// location in a structure usable for text selection and other reporting.
        /// </summary>
        public List<ErrorLocation> ErrorLocations { get { return _scanner.ErrorLocations; } }

        /// <summary>
        /// Results of rule checks from the session (lifespan 
        /// of the parser or after last call to 'Clean results;')
        /// </summary>
        public RuleCheckResultsManager RuleChecks { get { return _parser.RuleChecks; } }

        /// <summary>
        /// Messages go to the Console command line normally but 
        /// you can use this property to define optional output 
        /// where messages should go. 
        /// </summary>
        public System.IO.TextWriter Output { get { return _parser.Output; } set { _parser.Output = value; } }

        /// <summary>
        /// Model on which the parser operates
        /// </summary>
        public IModel Model { get { return _parser.Model; } }

        /// <summary>
        /// Variables which are result of the parsing process. 
        /// It can be either list of selected objects or new objects assigned to the variables.
        /// </summary>
        public XbimVariables Results { get { return _parser.Variables; } }

        /// <summary>
        /// Constructor which takes a existing model as an argument. 
        /// You can also close and open any model from the script.
        /// </summary>
        /// <param name="model">Model which shuldbe used for the script execution</param>
        public XbimQueryParser(IModel model)
        {
            Init(model);
        }

        /// <summary>
        /// Parameterless constructor of the class. 
        /// Default empty model is created which can be used 
        /// or you can open other model from the script.
        /// </summary>
        public XbimQueryParser()
        {
            Init(null);
        }

        private void Init(IModel model)
        {
            if (model == null)
                model = IModel.CreateTemporaryModel();
            _scanner = new Scanner();
            _parser = new Parser(_scanner, model);
            _parser.OnModelChanged += delegate(object sender, ModelChangedEventArgs e)
            {
                ModelChanged(e.NewModel);
            };
            _parser.OnFileReportCreated += delegate(object sender, FileReportCreatedEventArgs e)
            {
                FileReportCreated(e.FilePath);
            };
        }

        /// <summary>
        /// Set source for scanning and parsing
        /// </summary>
        /// <param name="source">source to be used</param>
        public void SetSource(string source)
        {
            ClearSources();
            _strSource = source;
        }

        /// <summary>
        /// Set source for scanning and parsing
        /// </summary>
        /// <param name="source">source to be used</param>
        public void SetSource(System.IO.Stream source)
        {
            ClearSources();
            _streamSource = source;
        }

        /// <summary>
        /// Set source for scanning and parsing
        /// </summary>
        /// <param name="source">source to be used</param>
        public void SetSource(IList<string> source)
        {
            ClearSources();
            _strListSource = source;
        }

        /// <summary>
        /// Performs only scan of the source and returns list of string 
        /// representation of Tokens. This is mainly for debugging.
        /// </summary>
        /// <returns>List of string representation of tokens</returns>
        public IEnumerable<string> ScanOnly()
        {
            ResetSource();
            List<string> result = new List<string>();
            int val = _scanner.yylex();
            while (val != (int)Tokens.EOF)
            {
                string name = val >= 60 ? Enum.GetName(typeof(Tokens), val) : ((char)val).ToString();
                result.Add(name);
                val = _scanner.yylex();
            }
            return result;
        }

        /// <summary>
        /// The main function used to perform parsing of the query. 
        /// Returns false only if something serious happen during
        /// parsing process. However it is quite possible that some errors occured. 
        /// So, make sure to check Errors if there are any.
        /// </summary>
        /// <returns>False if parsing failed, true otherwise.</returns>
        public bool Parse()
        {
            //no protection in debug mode so that it is easier to find problems
#if DEBUG
            ResetSource();
            var res = _parser.Parse();
            ScriptParsed();
            return res;
#else
            try
            {
                ResetSource();
                var res = _parser.Parse();
                ScriptParsed();
                return res;
            }
            catch (Exception)
            {
                //report error using standard error log of the scanner
                _scanner.Errors.Add("General parser error.");
                return false;
            }
#endif
        }

        /// <summary>
        /// The main function used to perform parsing of the query. 
        /// Returns false only if something serious happen during
        /// parsing process. However it is quite possible that some errors occured. 
        /// So, make sure to check Errors if there are any.
        /// </summary>
        /// <returns>False if parsing failed, true otherwise.</returns>
        public bool Parse(string source)
        {
            SetSource(source);
            return Parse();
        }

        /// <summary>
        /// The main function used to perform parsing of the query. 
        /// Returns false only if something serious happen during
        /// parsing process. However it is quite possible that some errors occured. 
        /// So, make sure to check Errors if there are any.
        /// </summary>
        /// <returns>False if parsing failed, true otherwise.</returns>
        public bool Parse(System.IO.Stream source)
        {
            SetSource(source);
            return Parse();
        }

        /// <summary>
        /// The main function used to perform parsing of the query. 
        /// Returns false only if something serious happen during
        /// parsing process. However it is quite possible that some errors occured. 
        /// So, make sure to check Errors if there are any.
        /// </summary>
        /// <returns>False if parsing failed, true otherwise.</returns>
        public bool Parse(IList<string> source)
        {
            SetSource(source);
            return Parse();
        }

        /// <summary>
        /// Source is available untill new source is defined. 
        /// So it is possible to perform scanning or parsing with the 
        /// source many times. Be carefull as side effects like data 
        /// creation will be persisted over the repeated execution.
        /// </summary>
        private void ResetSource()
        {
            if (_streamSource != null && _strListSource != null && _strSource != null) throw new Exception("Only one source can be valid.");
            if (_streamSource == null && _strListSource == null && _strSource == null) throw new Exception("One source must be valid.");
            if (_streamSource != null) _scanner.SetSource(_streamSource);
            if (_strListSource != null) _scanner.SetSource(_strListSource);
            if (_strSource != null) _scanner.SetSource(_strSource, 0);

            //reset error log with new source
            _scanner.Errors.Clear();
        }

        /// <summary>
        /// Helper function to clear sources available before new one is set.
        /// </summary>
        private void ClearSources() 
        {
            _strSource = null;
            _strListSource = null;
            _streamSource = null;
            //_scanner = new Scanner();
        }

        /// <summary>
        /// Event fired when model changes (closed or open)
        /// </summary>
        public event ModelChangedHandler OnModelChanged;
        private void ModelChanged(IModel newModel)
        {
            if (OnModelChanged != null)
                OnModelChanged(this, new ModelChangedEventArgs(newModel));
        }

        /// <summary>
        /// Event is fired when script is parsed
        /// </summary>
        public event ScriptParsedHandler OnScriptParsed;
        private void ScriptParsed()
        {
            if (OnScriptParsed != null)
                OnScriptParsed(this, new ScriptParsedEventArgs());
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

    }

    public delegate void ScriptParsedHandler(object sender, ScriptParsedEventArgs e);

    public class ScriptParsedEventArgs : EventArgs
    {
    }
}
