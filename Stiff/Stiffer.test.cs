using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace Stiff
{
    [TestFixture]
    public partial class Stiffer
    {
        [Test]
        public void シングルトン()
        {
            var st1 = Stiffer.GetInstance();
            Assert.True(st1 != null);

            var st2 = Stiffer.GetInstance();
            Assert.True(st1.Equals(st2));

            st1.Dispose();
            st2.Dispose();
        }


        [Test]
        public void アプリケーション起動()
        {
            //Assert.True(1 == 1);
            var st = Stiffer.GetInstance();

            st.CreateApplication();
            {
                Assert.True(st._app != null, "１回目");
                Assert.True(st._app.DisplayAlerts == false, "１回目 画面表示設定");

                var ap = st._app;
                st.CreateApplication();
                Assert.True(ap.Equals(st._app), "２回目");

                Marshal.ReleaseComObject(ap);
                Marshal.ReleaseComObject(st._app);
                ap = null;
                st._app = null;
            }
            st.Dispose();
            return;
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException))]
        public void ワークブックオープン_アプリ未起動()
        {
            var st = Stiffer.GetInstance();
            {
                Assert.True(st.OpenBook(@"c:\hoge.xls") == null);
            }
            st.Dispose();
        }

        [Test]
        [ExpectedException(typeof(System.Runtime.InteropServices.COMException))]
        public void ワークブックオープン_該当なし()
        {
            var st = Stiffer.GetInstance();
            st.CreateApplication();
            {
                Assert.True(st.OpenBook(@"c:\hoge.xls") == null);
            }
            st.Dispose();
        }

        [Test]
        public void ワークブックオープン_該当あり()
        {
            var st = Stiffer.GetInstance();
            st.CreateApplication();
            {
                var cd = System.IO.Directory.GetCurrentDirectory();
                Excel.Workbook oBook = null;

                try
                {
                    oBook = st.OpenBook(cd + @"\TestBook.xlsx");
                    Assert.True(oBook != null, "ファイルオープン");

                    var filename = oBook.FullName.ToString().ToUpper();
                    Assert.True(filename == (cd + @"\TestBook.xlsx").ToUpper(), "ファイルパス");
                }
                finally
                {
                    if (oBook != null)
                        Marshal.ReleaseComObject(oBook);
                    oBook = null;
                }
            }
            st.Dispose();
        }


    }
}
