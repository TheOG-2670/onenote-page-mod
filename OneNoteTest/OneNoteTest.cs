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

    public class Page
    {
        private string pageId;
        private XElement page, pageTitle, pageBody;

        public Page(string id)
        {
            pageId = id;
        }

        public string Id
        {
            get { return pageId; }
        }

        public XElement Title
        {
            set { pageTitle = value; }
            get { return pageTitle; }
        }

        public XElement Body
        {
            set { pageBody = value; }
            get { return pageBody; }
        }

        public void GetPageElements()
        {
            string pageXml;
            Utils.AppInstance.GetPageContent(pageId, out pageXml, PageInfo.piAll);

            XDocument pageDoc = XDocument.Parse(pageXml);
            page = pageDoc.Descendants(Utils.GetNameSpace() + "Page").First();
            XElement pageTitle_ = page.Descendants(Utils.GetNameSpace() + "Title").First();
            XElement pageOutline_ = page.Descendants(Utils.GetNameSpace() + "Outline").First();
            XElement pageBodyText_ = pageOutline_.Descendants(Utils.GetNameSpace() + "T").First();
            XElement pageTitleOE_ = pageTitle_.Descendants(Utils.GetNameSpace() + "OE").First();
            XElement pageTitleText_ = pageTitleOE_.Descendants(Utils.GetNameSpace() + "T").First();

            pageTitle = pageTitleText_;
            pageBody = pageBodyText_;
        }

        public void UpdateTitle(string newPageTitle)
        {
            if (!string.IsNullOrEmpty(newPageTitle))
            {
                pageTitle.Value=newPageTitle;
                Utils.AppInstance.UpdatePageContent(page.ToString());
            }
        }
    }

    internal class OneNoteTest
    {
        private static XNamespace xNameSpace;

        static void Main(string[] args)
        {
            xNameSpace=Utils.GetNameSpace();

            string notebookId = Utils.GetXmlObjectId(null, HierarchyScope.hsNotebooks, args[0]);
            //PrintElements(null, HierarchyScope.hsNotebooks, "name");

            string sectionId= Utils.GetXmlObjectId(notebookId, HierarchyScope.hsSections, args[1]);
            string pageId = Utils.GetXmlObjectId(sectionId, HierarchyScope.hsPages, args[2]);

            //Page p = new Page(pageId);
            //p.SetPageElements(pageId);
            //Utils.GetProperties(p);
        }



        //print the value of a specified attribute (such as name) for all elements
        private static void PrintElements(string parent, HierarchyScope scope, string attributeName)
        {
            Console.WriteLine("\n=====\nAll element names:\n=====");
            foreach (XElement element in Utils.GetXmlObjectElements(parent, scope))
            {
                Console.WriteLine(element.Attribute(attributeName).Value);
            }
        }
    }
}
