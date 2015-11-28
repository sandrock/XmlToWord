
namespace XmlToWord
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;

    class Context
    {
        private List<ErrorMessage> errors=new List<ErrorMessage>();

        public Context()
        {
        }

        public List<ErrorMessage> Errors
        {
            get { return this.errors; }
            set { this.errors = value; }
        }

        public Encoding XmlEncoding { get; set; }

        public XDocument Xml { get; set; }

        public MemoryStream WordStream { get; set; }

        public WordprocessingDocument Word { get; set; }

        public MainDocumentPart WordPart { get; set; }

        public Body WordBody { get; set; }

        public string Title { get; set; }

        public int? ExitCode { get; set; }

        internal void AddError(string message)
        {
            this.Errors.Add(new ErrorMessage(message));
        }

        internal void AddError(string message, string detail)
        {
            this.Errors.Add(new ErrorMessage(message, detail));
        }
    }

    class ErrorMessage
    {
        public ErrorMessage(string message)
        {
            this.Message = message;
        }

        public ErrorMessage(string message, string detail)
        {
            this.Message = message;
            this.Detail = detail;
        }

        public string Message { get; set; }

        public string Detail { get; set; }
    }

    class ItemPath
    {
        public string Path { get; set; }
        public string Attribute { get; set; }
        public ItemStyle Style { get; set; }
        public string Text { get; set; }
    }

    enum ItemStyle
    {
        Paragraph,
        Heading1,
        Heading2,
        Heading3,
    }
}
