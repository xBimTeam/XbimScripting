using System;
using System.Linq;
using Xbim.Ifc2x3.ProductExtension;
using Xbim.Ifc2x3.Extensions;
using System.IO;
using Xbim.Ifc2x3.ExternalReferenceResource;
using Xbim.Common;

namespace Xbim.Script
{
    internal class ClassificationCreator
    {
        private IModel _model;

        public void CreateSystem(IModel model, string classificationName)
        {
            //set model in which the systems are to be created
            _model = model;
            string data = null;

            //try to get data from resources

            //get list of classifications available
            var dir = Directory.EnumerateFiles("Classifications");
            string clasPath = null;
            foreach (var f in dir)
            {
                if (Path.GetFileNameWithoutExtension(f).ToLower() == classificationName.ToLower())
                {
                    clasPath = f;
                    break;
                }
            }
            if (clasPath == null)
                throw new ArgumentException("Specified classification doesn't exist");
            var clasName = Path.GetFileNameWithoutExtension(clasPath);


            //get data
            data = File.ReadAllText(clasPath);


            CsvLineParser parser = null;
            if (clasName.Contains("NRM"))
            {
                parser = new CsvLineParser(';');
            }
            else
                parser = new CsvLineParser();


            //create classification source
            var source = model.Instances.New<IfcClassification>(c =>
            {
                c.Source = clasName;
                c.Edition = "Default edition";
                c.Name = clasName;
            });
            ParseCSV(data, parser, source);
        }

        private void ParseCSV(string data, CsvLineParser parser, IfcClassification source)
        {
            TextReader csvReader = new StringReader(data);

            //header line
            string line = csvReader.ReadLine();
            CsvLineParser lineParser = parser;
            lineParser.ParseHeader(line);

            //get first line of data
            line = csvReader.ReadLine();

            while (line != null)
            {
                //parse line
                Line parsedLine = lineParser.ParseLine(line);

                //create IFC object
                IfcSystem system = GetOrCreateSystem(parsedLine.Code, parsedLine.Description);
                var classification = GetOrCreateClassificationReference(parsedLine.Code, parsedLine.Description, source);

                //set up hierarchy
                if (system == null) continue;
                IfcSystem parentSystem = GetOrCreateSystem(parsedLine.ParentCode);
                if (parentSystem != null) parentSystem.AddObjectToGroup(system);

                //read new line to be processed in the next step
                line = csvReader.ReadLine();
            }
        }


        private IfcSystem GetOrCreateSystem(string name, string description = null)
        {
            if (name == null) return null;

            IfcSystem system = _model.Instances.Where<IfcSystem>(s => s.Name == name).FirstOrDefault();
            if (system == null)
                system = _model.Instances.New<IfcSystem>(s =>
                {
                    s.Name = name;
                });
            if (description != null) system.Description = description;
            return system;
        }

        private IfcClassificationReference GetOrCreateClassificationReference(string code, string name, IfcClassification source)
        {
            if (code == null) return null;

            IfcClassificationReference classification = _model.Instances.Where<IfcClassificationReference>(s => s.ItemReference == code).FirstOrDefault();
            if (classification == null)
                classification = _model.Instances.New<IfcClassificationReference>(s =>
                {
                    s.ItemReference = code;
                });
            if (name != null) classification.Name = name;
            if (source != null) classification.ReferencedSource = source;

            return classification;
        }

        private class CsvLineParser
        {
            //settings
            private char _separator = ',';
            private string _CodeFieldName = "Code";
            private string _DescriptionFieldName = "Description";
            private string _ParentCodeFieldName = "Parent";

            //header used to parse the file
            private string[] header;
            //get index
            private int codeIndex;
            private int descriptionIndex;
            private int parentCodeIndex;

            public CsvLineParser(char separator = ',', string codeField = "Code", string descriptionField = "Description", string parentField = "Parent")
            {
                _separator = separator;
                _CodeFieldName = codeField;
                _DescriptionFieldName = descriptionField;
                _ParentCodeFieldName = parentField;
            }

            public void ParseHeader(string headerLine)
            {
                //create header of the file
                string line = headerLine.ToLower();
                header = line.Split(_separator);
                for (int i = 0; i < header.Count(); i++)
                    header[i] = header[i].Trim(' ', '"');

                //get indexes of the fields
                codeIndex = header.ToList().IndexOf(_CodeFieldName.ToLower());
                descriptionIndex = header.ToList().IndexOf(_DescriptionFieldName.ToLower());
                parentCodeIndex = header.ToList().IndexOf(_ParentCodeFieldName.ToLower());

                if (codeIndex < 0) throw new Exception("File is either not CSV file or it doesn't comply to the predefined structure.");
            }

            public Line ParseLine(string line)
            {
                Line result = new Line();
                string[] fields = line.Split(_separator);
                //trim all fields
                for (int i = 0; i < fields.Count(); i++)
                    fields[i] = fields[i].Trim(' ', '"');

                //get data
                int colNum = fields.Count();
                if (codeIndex >= 0 && colNum >= codeIndex + 1) result.Code = fields[codeIndex];
                if (descriptionIndex >= 0 && colNum >= descriptionIndex + 1) result.Description = fields[descriptionIndex];
                if (parentCodeIndex >= 0 && colNum >= parentCodeIndex + 1) result.ParentCode = fields[parentCodeIndex];

                return result;
            }
        }

        private struct Line
        {
            public string Code;
            public string Description;
            public string ParentCode;
        }
    }
}
