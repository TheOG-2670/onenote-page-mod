using Microsoft.Office.Interop.OneNote;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteTest
{
    public class Page
    {
        private string pageId;
        private XElement page, pageTitle;

        public Page(string id)
        {
            pageId = id;
            GetPageElements();
        }

        public string Id
        {
            get { return pageId; }
        }

        public XElement Title
        {
            get { return pageTitle; }
        }

        internal void GetPageElements()
        {
            Utils.AppInstance.GetPageContent(pageId, out string pageXml, PageInfo.piAll);

            XDocument pageDoc = XDocument.Parse(pageXml);
            page = pageDoc.Descendants(Utils.GetNameSpace() + "Page").First();
            XElement pageTitle_ = page.Descendants(Utils.GetNameSpace() + "Title").First();
            XElement pageTitleOE_ = pageTitle_.Descendants(Utils.GetNameSpace() + "OE").First();
            XElement pageTitleText_ = pageTitleOE_.Descendants(Utils.GetNameSpace() + "T").First();
            if(string.IsNullOrEmpty(pageTitleText_.Value))
            {
                pageTitleText_= (XElement)pageTitleOE_.Descendants(Utils.GetNameSpace() + "T").First().NextNode;
            }

            pageTitle = pageTitleText_;
        }

        public void UpdateTitle(string newPageTitle)
        {
            if (!string.IsNullOrEmpty(newPageTitle))
            {
                pageTitle.Value = newPageTitle;
                Utils.AppInstance.UpdatePageContent(page.ToString());
            }
        }
    }
}
