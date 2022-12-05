using Microsoft.Office.Interop.OneNote;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteTest
{
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
                pageTitle.Value = newPageTitle;
                Utils.AppInstance.UpdatePageContent(page.ToString());
            }
        }
    }
}
