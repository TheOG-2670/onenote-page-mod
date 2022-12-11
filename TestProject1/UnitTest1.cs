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
            DotNetEnv.Env.TraversePath().Load();

            _application = new Application();
            Utils.AppInstance = _application;

            notebookId = Utils.GetXmlObjectId(null, HierarchyScope.hsNotebooks, Environment.GetEnvironmentVariable("NOTEBOOK_TITLE"));
            sectionId = Utils.GetXmlObjectId(notebookId, HierarchyScope.hsSections, Environment.GetEnvironmentVariable("SECTION_TITLE"));
            pageId = Utils.GetXmlObjectId(sectionId, HierarchyScope.hsPages, Environment.GetEnvironmentVariable("CURRENT_PAGE_TITLE"));
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
            Assert.NotEmpty(p.Title.Value);

            string? expectedString = Environment.GetEnvironmentVariable("PAGE_BODY");
            expectedString?.ToList().ForEach(word =>
            {
                Assert.Contains(word, p.Body.Value);
            });
        }

        [Fact]
        public void CheckSameTitleRejected()
        {

            Page p = new Page(pageId);
            p.GetPageElements();

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
            p.GetPageElements();

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

        }

        ~UnitTest1()
        {
            ReleaseAppInstance();
        }
    }
}