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

            string notebookId = GetXmlObjectId(null, HierarchyScope.hsNotebooks, args[0]);
            PrintElements(null, HierarchyScope.hsNotebooks, "name");

            string sectionId= GetXmlObjectId(notebookId, HierarchyScope.hsSections, args[1]);
            string pageId = GetXmlObjectId(sectionId, HierarchyScope.hsPages, args[2]);
        }

        //get notebook's xml namespace used for traversing the notebook tree and searching for nodes (namespace + nodeName)
        private static void GetNameSpace()
        {
            string notebookXml;

            app.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
            XDocument doc = XDocument.Parse(notebookXml);
            xNameSpace = doc.Root.Name.Namespace;
        }

        //returns a list of xml elements from the hierarchy of a notebook, section, or page
        private static List<XElement> GetXmlObjectElements(string parent, HierarchyScope scope)
        {
            string xml;
            string nodeType = null;
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
            return doc.Descendants(xNameSpace + nodeType).ToList();
        }

        //search list of xml elements of an object (notebook, section, page) and return its ID based on the name provided
        private static string GetXmlObjectId(string parent, HierarchyScope scope, string objectName)
        {
            foreach (XElement element in GetXmlObjectElements(parent, scope))
            {
                if(element.Attribute("name").Value == objectName)
                {
                    return element.Attribute("ID").Value;
                }
            }
            return null;
        }

        //print the value of a specified attribute (such as name) for all elements
        private static void PrintElements(string parent, HierarchyScope scope, string attributeName)
        {
            Console.WriteLine("\n=====\nAll element names:\n=====");
            foreach (XElement element in GetXmlObjectElements(parent, scope))
            {
                Console.WriteLine(element.Attribute(attributeName).Value);
            }
        }
    }
}
