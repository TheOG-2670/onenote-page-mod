using Microsoft.Office.Interop.OneNote;
using System;

namespace OneNoteTest
{
    public sealed class OneNoteSingleton
    {
        private static Application instance;
        private static readonly object padlock = new object();

        private OneNoteSingleton()
        {
            //
        }

        public static Application Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Application();
                }
                return instance;
            }
        }
    }
}
