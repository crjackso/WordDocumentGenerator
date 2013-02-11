using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using LinqToExcel;
using OpenXmlPowerTools;

namespace EmployeeWordTemplater
{
    class Program
    {
        const string FinalTemplateRepository = @"Documents";

        static void Main()
        {
            var persons = new ExcelLoader().Load();
            var generator = new DocumentGenerator();

            ClearDocumentDirectory();

            foreach(var person in persons)
            {
                var document = generator.GenerateDocument(person);
                SaveDocument(document);
            }

            Console.WriteLine("Document Generation Complete");
            Console.ReadLine();
        }

        private static void ClearDocumentDirectory()
        {
            if (!Directory.Exists(FinalTemplateRepository))
                Directory.CreateDirectory(FinalTemplateRepository);

            var di = new DirectoryInfo(FinalTemplateRepository);
            di.EnumerateFiles().ForEach(o => o.Delete());
        }

        private static void SaveDocument(DocxContractDocument document)
        {
            string filePath = Path.Combine(FinalTemplateRepository, document.FileName);

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                fs.Write(document.Document, 0, document.Document.Length);
        }
    }

    public class ExcelLoader
    {
        private const string FileName = @"Resources\Employees.xlsx";

        public IEnumerable<Employee> Load()
        {
            var excel = CreateExcelFactory();
            var persons = from person in excel.Worksheet<Employee>()
                                   select person;

            return persons;
        }

        private static ExcelQueryFactory CreateExcelFactory()
        {
            var excel = new ExcelQueryFactory(FileName);
            excel.AddMapping<Employee>(x => x.FirstName, "First Name");
            excel.AddMapping<Employee>(x => x.LastName, "Last Name");
            excel.AddMapping<Employee>(x => x.Salary, "Salary Raise");
            return excel;
        }
    }

    public class Employee
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public decimal Salary { get; set; }

        public string SalaryFormat
        {
            get { return String.Format("{0:C}", Salary); }
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", FirstName, LastName);
        }
    }

    public class TemplateDocument
    {
        private readonly string _employeeName;
        private const string FilePath = @"Resources\EmployeeTemplate.docx";

        public byte[] Template { get; set; }

        public TemplateDocument(string employeeName)
        {
            _employeeName = employeeName;

            using (var fileStream = File.OpenRead(FilePath))
            {
                Template = new byte[fileStream.Length];
                fileStream.Read(Template, 0, Convert.ToInt32(fileStream.Length));
            }
        }

        public virtual byte[] CreateCopy()
        {
            var document = new byte[Template.Length];
            Template.CopyTo(document, 0);
            return document;
        }

        public virtual string DocumentName
        {
            get { return String.Format("{0}-{1:yyyy-MM-dd}.docx", _employeeName, DateTime.Now); }
        }

        public virtual string MimeType
        {
            get { return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; }
        }
    }

    public class DocumentGenerator
    {
        private readonly FieldMerger _fieldMerger;
        private readonly Dictionary<string, FieldMapperFunction> _fieldMapping;

        private delegate IContentReplacer FieldMapperFunction(Employee employee);

        public DocumentGenerator()
        {
            _fieldMerger = new FieldMerger();
            _fieldMapping = CreateFieldMapping();
        }

        private Dictionary<string, FieldMapperFunction> CreateFieldMapping()
        {
            return new Dictionary<string, FieldMapperFunction>
                       {
                           {"EmployeeName", (person) => new SingleLineContentReplacer(person.ToString())},
                           {"SalaryRaise", (person) => new SingleLineContentReplacer(person.SalaryFormat)}
                       };
        }

        public DocxContractDocument GenerateDocument(Employee employee)
        {
            var template = new TemplateDocument(employee.ToString());
            var templateCopy = template.CreateCopy();
            var mergeFields = _fieldMerger.GetContentControlValues(templateCopy).ToList();

            var fieldValuesDictionary = new Dictionary<string, IContentReplacer>();

            foreach (var mergeField in mergeFields)
            {
                fieldValuesDictionary[mergeField.Key] = _fieldMapping[mergeField.Key](employee);
            }

            return new DocxContractDocument(_fieldMerger.Merge(templateCopy, fieldValuesDictionary), template.DocumentName);
        }
    }

    public class FieldMerger
    {
        public virtual byte[] Merge(byte[] templateCopy, Dictionary<string, IContentReplacer> contentReplacers)
        {
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(templateCopy, 0, templateCopy.Length);
                using (var wordProcessingDocument = WordprocessingDocument.Open(memoryStream, true))
                {
                    var contentControls = FindContentControls(wordProcessingDocument).ToList();
                    foreach (var keyValue in contentReplacers.Where(keyValue => keyValue.Value != null))
                    {
                        Replace(contentControls, keyValue.Key, keyValue.Value);
                    }

                    wordProcessingDocument.ContentParts().ForEach(p => p.PutXDocument());
                }
                return memoryStream.ToArray();
            }
        }

        private static void Replace(IEnumerable<WordDocumentContentControl> contentControls, string fieldKey, IContentReplacer contentReplacer)
        {
            contentControls.Where(mf => mf.Key == fieldKey).ForEach(mf => mf.ReplaceContent(contentReplacer));
        }

        public virtual IEnumerable<WordDocumentContentControl> GetContentControlValues(byte[] document)
        {
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(document, 0, document.Length);
                using (var wordprocessingDocument = WordprocessingDocument.Open(memoryStream, true))
                {
                    return FindContentControls(wordprocessingDocument);
                }
            }
        }

        public virtual IEnumerable<string> GetContentFields(byte[] document)
        {
            var fields = GetContentControlValues(document);
            return fields.Select(f => f.Key);
        }

        private static IEnumerable<WordDocumentContentControl> FindContentControls(WordprocessingDocument wordProcessingDocument)
        {
            var allSdtRuns = wordProcessingDocument.ContentParts().SelectMany(p => p.GetXDocument().Descendants(W.sdt)).ToList();

            var contentTextFields = WordDocumentContentControl.GetContentControls(allSdtRuns);

            return contentTextFields;
        }
    }

    public interface IContentReplacer
    {
        void ReplaceContent(XElement sdtContentElement);
    }

    public class SingleLineContentReplacer : IContentReplacer
    {
        private readonly string _text;

        public SingleLineContentReplacer(string text)
        {
            _text = text;
        }

        public void ReplaceContent(XElement sdtContentElement)
        {
            var texts = sdtContentElement.Descendants(W.t).ToList();

            if (texts.Count > 1)
                texts.Skip(1).ForEach(el => el.Remove());

            texts.First().SetValue(_text);
        }

        public override string ToString()
        {
            return _text;
        }
    }

    public class WordDocumentContentControl
    {

        private readonly string _key;
        private readonly XElement _sdtContentElement;

        public WordDocumentContentControl(string key, XElement sdtContentElement)
        {
            _key = key;
            _sdtContentElement = sdtContentElement;
        }

        public string Key
        {
            get { return _key; }
        }

        public void ReplaceContent(IContentReplacer contentReplacer)
        {
            contentReplacer.ReplaceContent(_sdtContentElement);
        }

        public string GetTextValue()
        {
            return _sdtContentElement.Descendants(W.t).Aggregate(new StringBuilder(), (sb, x) => sb.Append(x.Value), sb => sb.ToString());
        }

        public List<List<string>> GetTableValues()
        {
            return _sdtContentElement.Descendants(W.tr).Select(row => row.Descendants(W.t).Select(x => x.Value).ToList()).ToList();
        }

        public static IEnumerable<WordDocumentContentControl> GetContentControls(IEnumerable<XElement> runs)
        {
            return runs.Select(run =>
            {
                var key = run.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).First().Value;
                var sdtContentElement = run.Elements(W.sdtContent).First();
                return new WordDocumentContentControl(key, sdtContentElement);
            });

        }
    }

    public class DocxContractDocument
    {
        public string FileName { get; set; }
        public virtual byte[] Document { get; set; }

        public DocxContractDocument(byte[] document, string documentName)
        {
            Document = document;
            FileName = documentName;
        }

        public string MimeType
        {
            get { return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; }
        }

        public string Extension
        {
            get { return "docx"; }
        }
    }

    public static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> enumerable, Action<T> action)
        {
            foreach (var item in enumerable) action(item);
        }
    }
}
