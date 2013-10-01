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

        [Test]
        public void プロパティ_なし()
        {
            var st = Stiffer.GetInstance();
            st.CreateApplication();
            {
                var cd = System.IO.Directory.GetCurrentDirectory();
                Excel.Workbook oBook = null;

                try
                {
                    oBook = st.OpenBook(cd + @"\TestBook.xlsx");
                    string value = st.GetBuiltinProperty(oBook, "hoge");
                    Assert.True( value == "");
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

        [Test]
        public void プロパティ_有り()
        {
            var st = Stiffer.GetInstance();
            st.CreateApplication();
            try
            {
                var cd = System.IO.Directory.GetCurrentDirectory();
                Excel.Workbook oBook = null;

                try
                {
                    oBook = st.OpenBook(cd + @"\TestBook.xlsx");
                    string value = st.GetBuiltinProperty(oBook, "Author");
                    Assert.True(value == "小林礼明", "Author");
                    Assert.True(st.GetBuiltinProperty(oBook, "Title") == "テスト用ブック", "タイトル");
                    Assert.True(st.GetBuiltinProperty(oBook, "Subject") == "Stiff", "サブタイトル");
                    Assert.True(st.GetBuiltinProperty(oBook, "Company") == "個人", "会社");
                    Assert.True(st.GetBuiltinProperty(oBook, "Manager") == "わし", "管理者" + st.GetBuiltinProperty(oBook, "Manager"));
                    Assert.True(st.GetBuiltinProperty(oBook, "Last Save Time") == "2013/10/01 23:03:27", "前回保存日時");
                }
                finally
                {
                    if (oBook != null)
                        Marshal.ReleaseComObject(oBook);
                    oBook = null;
                }
            }
            finally
            {
                st.Dispose();
            }
        }


        [Test]
        [ExpectedException(typeof(System.Runtime.InteropServices.COMException))]
        public void ブック情報取得_該当なし()
        {
            var st = Stiffer.GetInstance();
            try
            {
                var cd = System.IO.Directory.GetCurrentDirectory();

                var info = st.GetInformations(cd + @"\Hoge.xlsx");
            }
            finally
            {
                st.Dispose();
            }
        }

        [Test]
        public void ブック情報取得_該当あり()
        {
            var st = Stiffer.GetInstance();
            try
            {
                var cd = System.IO.Directory.GetCurrentDirectory();

                var info = st.GetInformations(cd + @"\TestBook.xlsx");

                Assert.True(info != null, "ヌルチェック");
                Assert.True(info.Author         == "小林礼明", "Author");
                Assert.True(info.Title          == "テスト用ブック", "タイトル");
                Assert.True(info.Subject        == "Stiff", "サブタイトル");
                Assert.True(info.Company        == "個人", "会社");
                Assert.True(info.Manager        == "わし", "管理者");
                Assert.True(info.LastSaveTime   == "2013/10/01 23:03:27", "前回保存日時");
            }
            finally
            {
                st.Dispose();
            }
        }

        [Test]
        //[ExpectedException(typeof(System.Runtime.InteropServices.COMException))]
        public void ブック情報取得_Excel外ファイル()
        {
            var st = Stiffer.GetInstance();
            try
            {
                var cd = System.IO.Directory.GetCurrentDirectory();

                var info = st.GetInformations(cd + @"\Stiff.exe");
                Assert.True( info.LastSaveTime == "" );
            }
            finally
            {
                st.Dispose();
            }
        }

        [Test]
        public void プロパティ_変更()
        {
            var st = Stiffer.GetInstance();
            st.CreateApplication();
            try
            {
                var cd = System.IO.Directory.GetCurrentDirectory();
                Excel.Workbook oBook = null;

                try
                {
                    BookInfo info = this.GetInformations(cd + @"\TestBook.xlsx");

                    info.Author  = "kobayashi";
                    info.Title   = "Modified";
                    info.Subject = "Tool";
                    info.Company = "The Man";
                    info.Manager = "Myself";

                    // 書き換え
                    oBook = st.OpenBook(cd + @"\TestBook.xlsx");
                    this.SetInformations(oBook, info);

                    // 確認
                    Assert.True(st.GetBuiltinProperty(oBook, "Author"        ) == "kobayashi", "Author");
                    Assert.True(st.GetBuiltinProperty(oBook, "Title"         ) == "Modified", "タイトル");
                    Assert.True(st.GetBuiltinProperty(oBook, "Subject"       ) == "Tool", "サブタイトル");
                    Assert.True(st.GetBuiltinProperty(oBook, "Company"       ) == "The Man", "会社");
                    Assert.True(st.GetBuiltinProperty(oBook, "Manager"       ) == "Myself", "管理者" + st.GetBuiltinProperty(oBook, "Manager"));
                    Assert.True(st.GetBuiltinProperty(oBook, "Last Save Time") == "2013/10/01 23:03:27", "前回保存日時");
                }
                finally
                {
                    if (oBook != null)
                        Marshal.ReleaseComObject(oBook);
                    oBook = null;
                }
            }
            finally
            {
                st.Dispose();
            }
        }
    }
}
