using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Linq;
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

            string notebookId = GetXmlObjectId(null, HierarchyScope.hsNotebooks, args[0]);
            Console.WriteLine(notebookId);

            string sectionId= GetXmlObjectId(notebookId, HierarchyScope.hsSections, args[1]);
            Console.WriteLine(sectionId);

            string pageId = GetXmlObjectId(sectionId, HierarchyScope.hsPages, args[2]);
            Console.WriteLine(pageId);
        }

        //get notebook's xml namespace used for traversing the notebook tree and searching for nodes (namespace + nodeName)
        private static void GetNameSpace()
        {
            string notebookXml;

            app.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
            XDocument doc = XDocument.Parse(notebookXml);
            xNameSpace = doc.Root.Name.Namespace;
        }

        //traverse the xml tree for a given object (notebook, section, page) and return its ID
        private static string GetXmlObjectId(string parent, HierarchyScope scope, string objectName)
        {
            string xml;
            string nodeType=null;

            app.GetHierarchy(parent, scope, out xml);

            switch (scope)
            {
                case HierarchyScope.hsNotebooks:
                    nodeType = "Notebook";
                    break;
                case HierarchyScope.hsSections:
                    nodeType = "Section";
                    break;
                case HierarchyScope.hsPages:
                    nodeType = "Page";
                    break;
            }
            
            XDocument doc = XDocument.Parse(xml);
            List<XElement> docElements = doc.Descendants(xNameSpace + nodeType).ToList();
            foreach (XElement element in docElements)
            {
                if(element.Attribute("name").Value == objectName)
                {
                    return element.Attribute("ID").Value;
                }
            }

            return null;
        }
    }
}
