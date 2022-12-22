using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace TestProject1
{
    public class UnitTest1
    {
        private static Application? _application;
        private static string notebookId="", sectionId="", pageId="";

        internal static void ReleaseAppInstance()
        {
            if(OneNoteSingleton.Instance != null && _application != null)
            {
                Marshal.ReleaseComObject(OneNoteSingleton.Instance);
                _application = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            
        }

        public UnitTest1()
        {
            DotNetEnv.Env.TraversePath().Load();

            _application = new Application();

            notebookId = Utils.GetXmlObjectId(null, HierarchyScope.hsNotebooks, Environment.GetEnvironmentVariable("NOTEBOOK_TITLE"));
            sectionId = Utils.GetXmlObjectId(notebookId, HierarchyScope.hsSections, Environment.GetEnvironmentVariable("SECTION_TITLE"));
            pageId = Utils.GetXmlObjectId(sectionId, HierarchyScope.hsPages, Environment.GetEnvironmentVariable("CURRENT_PAGE_TITLE"));
        }

        [Fact]
        public void CheckAppInstanceNotNull()
        {
            Assert.NotNull(OneNoteSingleton.Instance);   
        }

        [Fact]
        public void CheckIdsNotEmpty()
        {

            Assert.NotEmpty(notebookId);
            Assert.NotEmpty(sectionId);
            Assert.NotEmpty(pageId);
        }

        [Fact]
        public void CheckChangePageTitle()
        {

            Page p = new Page(pageId);

            string? newTitle = Environment.GetEnvironmentVariable("NEW_PAGE_TITLE");
            p.UpdateTitle(newTitle);
            Assert.Equal(newTitle, p.Title.Value);
            try
            {
                p.UpdateTitle(Environment.GetEnvironmentVariable("CURRENT_PAGE_TITLE"));
            }
            catch(Exception ex)
            {
                Assert.NotNull(ex.Message);
            }
            Assert.Equal(Environment.GetEnvironmentVariable("CURRENT_PAGE_TITLE"), p.Title.Value);
        }

        [Fact]
        public void CheckDifferentTitleAccepted()
        {

            Page p = new Page(pageId);

            if (!Environment.GetEnvironmentVariable("CURRENT_PAGE_TITLE").Equals(p.Title.Value))
            {

                try
                {
                    p.UpdateTitle(Environment.GetEnvironmentVariable("NEW_PAGE_TITLE"));
                    Assert.Equal(Environment.GetEnvironmentVariable("NEW_PAGE_TITLE"), p.Title.Value);
                }
                catch (Exception ex)
                {
                    //
                }
            }
            string? newTitle = Environment.GetEnvironmentVariable("NEW_PAGE_TITLE");
            p.UpdateTitle(newTitle);
            Assert.Equal(newTitle, p.Title.Value);
        }

        [Fact]
        public void CheckAccessMultiplePages()
        {
            List<Page> pages = new List<Page>();
            List<XElement> xPages = Utils.GetXmlObjectElements(sectionId, HierarchyScope.hsPages);

            Assert.NotEmpty(xPages);

            foreach(XElement p in xPages)
            {
                string? id = p.Attribute("ID")?.Value;
                if (string.IsNullOrEmpty(id))
                {
                    break;
                }
                pages.Add(new Page(id));
            }

            pages.ForEach(p => Assert.NotEmpty(p.Title.Value));
        }

        ~UnitTest1()
        {
            ReleaseAppInstance();
        }
    }
}