using System.Xml.Linq;

namespace TemplateEngine.Docx
{
    internal static class W14
    {
        public static XNamespace w14 =
            "http://schemas.microsoft.com/office/word/2010/wordml";

        public static XName checkbox = w14 + "checkbox";
        public static XName checkedState = w14 + "checkedState";
        public static XName uncheckedState = w14 + "uncheckedState";

        public static XName val = w14 + "val";
        public static XName font = w14 + "font";

    }
}