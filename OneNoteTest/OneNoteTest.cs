using Microsoft.Office.Interop.OneNote;
using System;
using System.Xml.Linq;

namespace OneNoteTest
{
    internal class OneNoteTest
    {
        private static XNamespace xNameSpace;

        static void Main(string[] args)
        {
            xNameSpace=Utils.GetNameSpace();

            string notebookId = Utils.GetXmlObjectId(null, HierarchyScope.hsNotebooks, args[0]);
            string sectionId= Utils.GetXmlObjectId(notebookId, HierarchyScope.hsSections, args[1]);
            string pageId = Utils.GetXmlObjectId(sectionId, HierarchyScope.hsPages, args[2]);

            try
            {
                Page p = new Page(pageId);
                p.UpdateTitle(args[3]);
                Console.WriteLine($"Existing title: {args[2]} \nUpdated title: {p.Title.Value}");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error occurred during title update: {ex.Message}");
            }
            
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
