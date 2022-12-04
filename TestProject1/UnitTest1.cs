using System.Runtime.InteropServices;

namespace TestProject1
{
    public class UnitTest1
    {
        private static Application? _application;
        private static string notebookId="", sectionId="", pageId="";

        internal static void ReleaseAppInstance()
        {
            if(Utils.AppInstance!=null && _application != null)
            {
                Marshal.ReleaseComObject(Utils.AppInstance);
                Utils.AppInstance = null;
                _application = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            
        }

        public UnitTest1()
        {
            _application = new Application();
            Utils.AppInstance = _application;
            notebookId = Utils.GetXmlObjectId(null, HierarchyScope.hsNotebooks, "My Notebook");
            sectionId = Utils.GetXmlObjectId(notebookId, HierarchyScope.hsSections, "Miscellaneous");
            pageId = Utils.GetXmlObjectId(sectionId, HierarchyScope.hsPages, "hello123hi");
        }

        [Fact]
        public void CheckAppInstanceNotNull()
        {
            Assert.NotNull(Utils.AppInstance);   
        }

        [Fact]
        public void CheckIdsNotEmpty()
        {

            Assert.NotEmpty(notebookId);
            Assert.NotEmpty(sectionId);
            Assert.NotEmpty(pageId);
        }

        [Fact]
        public void CheckPageInfoIsCorrect()
        {


            Page p = new Page(pageId);
            p.GetPageElements();

            Assert.NotNull(p.Id);
            Assert.Equal("hello123hi", p.Title.Value);

            string expectedString = "this is a test";
            expectedString.ToList().ForEach(word =>
            {
                Assert.Contains(word, p.Body.Value);
            });
        }

        [Fact]
        public void CheckChangePageTitle()
        {

            Page p = new Page(pageId);
            p.GetPageElements();

            string newTitle = "hello123";
            p.UpdateTitle(newTitle);
            Assert.Equal(newTitle, p.Title.Value);
        }

        ~UnitTest1()
        {
            ReleaseAppInstance();
        }
    }
}