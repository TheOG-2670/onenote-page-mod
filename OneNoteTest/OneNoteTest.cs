using Microsoft.Office.Interop.OneNote;
using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OneNoteTest
{
    internal class OneNoteTest
    {
        static Application app;
        static XNamespace xNameSpace;

        static void Main(string[] args)
        {
            app = new Application();

            GetNameSpace();
            Console.WriteLine(xNameSpace);
        }

        //get notebook's xml namespace used for traversing the notebook tree and searching for nodes (namespace + nodeName)
        private static void GetNameSpace()
        {
            string notebookXml;

            app.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
            XDocument doc = XDocument.Parse(notebookXml);
            xNameSpace = doc.Root.Name.Namespace;
        }
    }
}
