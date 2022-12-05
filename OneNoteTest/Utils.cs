using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteTest
{
    public class Utils
    {
        public static Application AppInstance
        {
            get;
            set;
        }


        //get notebook's xml namespace used for traversing the notebook tree and searching for nodes (namespace + nodeName)
        public static XNamespace GetNameSpace()
        {
            string notebookXml;

            AppInstance.GetHierarchy(null, HierarchyScope.hsNotebooks, out notebookXml);
            XDocument doc = XDocument.Parse(notebookXml);
            return doc.Root.Name.Namespace;
        }

        //returns a list of xml elements from the hierarchy of a notebook, section, or page
        public static List<XElement> GetXmlObjectElements(string parent, HierarchyScope scope)
        {
            string xml;
            string nodeType = null;
            AppInstance.GetHierarchy(parent, scope, out xml);

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
            return doc.Descendants(GetNameSpace() + nodeType).ToList();
        }

        //search list of xml elements of an object (notebook, section, page) and return its ID based on the name provided
        public static string GetXmlObjectId(string parent, HierarchyScope scope, string objectName)
        {
            foreach (XElement element in GetXmlObjectElements(parent, scope))
            {
                if (element.Attribute("name").Value == objectName)
                {
                    return element.Attribute("ID").Value;
                }
            }
            return null;
        }

        public static void GetProperties(object obj)
        {
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(obj))
            {
                Console.WriteLine($"{descriptor.Name}: {descriptor.GetValue(obj)}");
            }
        }
    }
}
