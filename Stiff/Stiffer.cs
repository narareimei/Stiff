using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Reflection;
using System.Diagnostics;

namespace Stiff
{
    public class BookInfo
    {
        public String Author { get; set; }
        public String Title { get; set; }
        public String SubTitle { get; set; }
        public String Manager { get; set; }
        public String Company { get; set; }

        public String FileName { get; set; }
        public String UpdateDate { get; set; }
    }

    public partial class Stiffer : IDisposable
    {
        /// <summary>
        /// シングルトンインスタンス
        /// </summary>
        private static Stiffer _instance;

        /// <summary>
        /// Excelアプリケーションインスタンス
        /// </summary>
        private Excel.Application _app;

        /// <summary>
        /// Dispose済みフラグ
        /// </summary>
        private bool _disposed;


        /// <summary>
        /// 静的コンストラクタ
        /// </summary>
        static Stiffer()
        {
            _instance = null;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Stiffer()
        {
            _app = null;
            _disposed = false;
        }

        /// <summary>
        ///インスタンス取得（シングルトン） 
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// インスタンスを１つにしたいわけでもない・・・
        /// </remarks>
        public static Stiffer GetInstance()
        {
            if(_instance == null) {
                _instance = new Stiffer();
            }
            else if (_instance._disposed == true)
            {
                _instance = new Stiffer();
            }
            return _instance;
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("B196B283-BAB4-101A-B69C-00AA00341D07")]
        public interface IProvideClassInfo
        {
            [return: MarshalAs(UnmanagedType.Interface)]
            UCOMITypeInfo GetClassInfo();
        }

        /// <summary>
        /// 
        /// </summary>
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
        public interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount();
            [PreserveSig]
            int GetTypeInfo([In] int index, [In] int lcid, [MarshalAs(UnmanagedType.Interface)] out UCOMITypeInfo pTypeInfo);
            [PreserveSig]
            int GetIDsOfNames();
            [PreserveSig]
            int Invoke();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comobj_"></param>
        /// <returns></returns>
        public static string GetComClassName(object comobj_)
        {
            if (Marshal.IsComObject(comobj_) == false)
            {
                return null;
            }
            IDispatch dispatch = (IDispatch)comobj_;

            UCOMITypeInfo typeinfo = null;
            if (typeinfo == null && comobj_ is IDispatch)
            {
                dispatch.GetTypeInfo(0, 0x409, out typeinfo);
            }
            if (typeinfo == null && comobj_ is IProvideClassInfo)
            {
                IProvideClassInfo provideclassinfo = (IProvideClassInfo)comobj_;
                typeinfo = provideclassinfo.GetClassInfo();
            }
            if (typeinfo != null)
            {
                string strName;
                string strDocString;
                int dwHelpContext;
                string strHelpFile;
                typeinfo.GetDocumentation(-1, out strName, out strDocString, out dwHelpContext, out strHelpFile);

                //{
                //    Object[] args = new Object[5];
                //    string[] rgsNames = new string[1];
                //    rgsNames[0] = "PrintNormal";

                //    uint LOCALE_SYSTEM_DEFAULT = 0x0800;
                //    uint lcid = LOCALE_SYSTEM_DEFAULT;
                //    int cNames = 1;
                //    int[] rgDispId = new int[1];
                //    args[0] = IntPtr.Zero;
                //    args[1] = rgsNames;
                //    args[2] = cNames;
                //    args[3] = lcid;
                //    args[4] = rgDispId;

                //    int [] pMemId = new int [100];
                //    typeinfo.GetIDsOfNames(rgsNames, cNames, pMemId);
                //}

                return strName;
            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        public void CreateApplication()
        {
            // 有効チェック
            if (_disposed)
                throw new ObjectDisposedException("Resource was disposed.");

            if (_app == null)
            {
                // Excelプロセスを毎回起動すると重いので一度だけにする
                this._app = new Excel.Application();
                this._app.DisplayAlerts = false;
            }
            return;
        }


        /// <summary>
        /// ワークブックを開く
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public Excel.Workbook OpenBook(string filename)
        {
            var             oBooks = this._app.Workbooks;
            Excel.Workbook  oBook  = null;

            try
            {
                oBook = oBooks.Open(
                              filename,  // オープンするExcelファイル名
                              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                              Type.Missing, // （省略可能）ReadOnly (True / False )
                              Type.Missing, // （省略可能）Format
                                            // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                                            // 5:なし / 6:引数 Delimiterで指定された文字
                              Type.Missing, // （省略可能）Password
                              Type.Missing, // （省略可能）WriteResPassword
                              Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                              Type.Missing, // （省略可能）Origin
                              Type.Missing, // （省略可能）Delimiter
                              Type.Missing, // （省略可能）Editable
                              Type.Missing, // （省略可能）Notify
                              Type.Missing, // （省略可能）Converter
                              Type.Missing, // （省略可能）AddToMru
                              Type.Missing, // （省略可能）Local
                              Type.Missing  // （省略可能）CorruptLoad
                          );
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(oBooks);
                oBooks = null;
            }
            return oBook;
        }

        /// <summary>
        /// Excelファイルのプロパティを取得する
        /// </summary>
        /// <param name="oBook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public string GetBuiltinProperty(Excel.Workbook oBook, string propertyName)
        {
            string strValue = "";

            try
            {
                object ps = oBook.BuiltinDocumentProperties;
                Type typeDocBuiltInProps = ps.GetType();

                //Get the Author property and display it.
                object oDocAuthorProp = typeDocBuiltInProps.InvokeMember("Item",
                                           BindingFlags.Default | BindingFlags.GetProperty,
                                           null, ps, new object[] { propertyName });

                Debug.Assert(oDocAuthorProp != null);
                Type typeDocAuthorProp = oDocAuthorProp.GetType();
                strValue = typeDocAuthorProp.InvokeMember("Value",
                                           BindingFlags.Default | BindingFlags.GetProperty,
                                           null, oDocAuthorProp, new object[] { }).ToString();

                Console.WriteLine(string.Format("The Excel file: '{0}' Last modified value = '{1}'", "", strValue));
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("The application failed with the following exception: '{0}' Stacktrace:'{1}'", ex.Message, ex.StackTrace));
            }
            finally
            {
            }
            return strValue;
        }


        /// <summary>
        /// Excelブックの各種情報取得
        /// </summary>
        public BookInfo GetInformations(string filename)
        {
            //Unification();

            Excel.Workbooks oBooks = null;
            Excel.Workbook oBook = null;
            Excel.Sheets oSheets = null;
            Excel.Worksheet oSheet = null;
            Excel.Range oCells = null;
            Excel.Range oRange = null;
            //
            Microsoft.Office.Core.DocumentProperties documentProperties = null;

            var info = new BookInfo();
            try
            {
                info.FileName = filename;

                // 有効チェック
                if (_disposed)
                    throw new ObjectDisposedException("Resource was disposed.");

                if (_app == null)
                {
                    // Excelプロセスを毎回起動すると重いので一度だけにする
                    this._app = new Excel.Application();
                    this._app.DisplayAlerts = false;
                }
                // ファイルオープン
                {
                    var books = this._app.Workbooks;
                    oBook = books.Open(
                                  filename,  // オープンするExcelファイル名
                                  Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                                  Type.Missing, // （省略可能）ReadOnly (True / False )
                                  Type.Missing, // （省略可能）Format
                        // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                        // 5:なし / 6:引数 Delimiterで指定された文字
                                  Type.Missing, // （省略可能）Password
                                  Type.Missing, // （省略可能）WriteResPassword
                                  Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                                  Type.Missing, // （省略可能）Origin
                                  Type.Missing, // （省略可能）Delimiter
                                  Type.Missing, // （省略可能）Editable
                                  Type.Missing, // （省略可能）Notify
                                  Type.Missing, // （省略可能）Converter
                                  Type.Missing, // （省略可能）AddToMru
                                  Type.Missing, // （省略可能）Local
                                  Type.Missing  // （省略可能）CorruptLoad
                              );
                    object ps = oBook.BuiltinDocumentProperties;
                    //{
                    //    Object[] args = new Object[5];
                    //    string[] rgsNames = new string[1];
                    //    rgsNames[0] = "PrintNormal";

                    //    uint LOCALE_SYSTEM_DEFAULT = 0x0800;
                    //    uint lcid = LOCALE_SYSTEM_DEFAULT;
                    //    int cNames = 1;
                    //    int[] rgDispId = new int[1];
                    //    args[0] = IntPtr.Zero;
                    //    args[1] = rgsNames;
                    //    args[2] = cNames;
                    //    args[3] = lcid;
                    //    args[4] = rgDispId;
                    //    Object result = ps.GetType().InvokeMember("GetIDsOfNames", BindingFlags.InvokeMethod, null, ps, args);
                    //}
                    GetComClassName(ps);


                    {
                        var strvalue = this.GetBuiltinProperty(oBook, "Title");
                        //try
                        //{
                        //    Type typeDocBuiltInProps = ps.GetType();
                        //    //Get the Author property and display it.
                        //    string strIndex = "Last Save Time";
                        //    string strValue;
                        //    object oDocAuthorProp = typeDocBuiltInProps.InvokeMember("Item",
                        //                               BindingFlags.Default |
                        //                               BindingFlags.GetProperty,
                        //                               null, ps,
                        //                               new object[] { strIndex });
                        //    Type typeDocAuthorProp = oDocAuthorProp.GetType();
                        //    strValue = typeDocAuthorProp.InvokeMember("Value",
                        //                               BindingFlags.Default |
                        //                               BindingFlags.GetProperty,
                        //                               null, oDocAuthorProp,
                        //                               new object[] { }).ToString();

                        //    Console.WriteLine(string.Format("The Excel file: '{0}' Last modified value = '{1}'", "", strValue));
                        //}
                        //catch (Exception ex)
                        //{
                        //    Console.WriteLine(string.Format("The application failed with the following exception: '{0}' Stacktrace:'{1}'", ex.Message, ex.StackTrace));
                        //}
                        //finally
                        //{
                        //}
                    }


                    documentProperties = (Microsoft.Office.Core.DocumentProperties)oBook.BuiltinDocumentProperties;
                    //foreach (var prop in documentProperties)
                    //{

                    //}


                    //object oBuiltInProps = oBook.BuiltinDocumentProperties;
                    ////Get the value of the Author property and display it
                    //string strValue = oBuiltInProps.Item("Author").Value;

                    DocumentProperty prop = documentProperties["Author"];
                    info.Author = (String)prop.Value;
                    Marshal.ReleaseComObject(prop);
                    prop = null;
                }

            }
            finally
            {
                if (documentProperties != null)
                {
                    Marshal.ReleaseComObject(documentProperties);
                }
                documentProperties = null;

                if (oRange != null)
                {
                    Marshal.ReleaseComObject(oRange);
                }
                oRange = null;

                if (oCells != null)
                {
                    Marshal.ReleaseComObject(oCells);
                }
                oCells = null;

                if (oSheet != null)
                {
                    Marshal.ReleaseComObject(oSheet);
                }
                oSheet = null;


                if (oSheets != null)
                {
                    Marshal.ReleaseComObject(oSheets);
                }
                oSheets = null;

                if (oBook != null)
                {
                    oBook.Close(false, filename, Type.Missing);
                    Marshal.ReleaseComObject(oBook);
                }
                oBook = null;

                if (oBooks != null)
                {
                    Marshal.ReleaseComObject(oBooks);
                }
                oBooks = null;
            }
            return null;
        }

        public void Unification()
        {
            // 有効チェック
            if (_disposed)
                throw new ObjectDisposedException("Resource was disposed.");


            // 一度お試しで
            var filename = @"C:\Users\Administrator\Dropbox\private\dotNet\Stiff\Stiff\TestBook.xlsx";
            if (_app == null)
            {
                // Excelプロセスを毎回起動すると重いので一度だけにする
                this._app = new Excel.Application();
                this._app.DisplayAlerts = false;
            }
            {
                var books = this._app.Workbooks;
                var oBook = books.Open(
                              filename,  // オープンするExcelファイル名
                              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                              Type.Missing, // （省略可能）ReadOnly (True / False )
                              Type.Missing, // （省略可能）Format
                                            // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                                            // 5:なし / 6:引数 Delimiterで指定された文字
                              Type.Missing, // （省略可能）Password
                              Type.Missing, // （省略可能）WriteResPassword
                              Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                              Type.Missing, // （省略可能）Origin
                              Type.Missing, // （省略可能）Delimiter
                              Type.Missing, // （省略可能）Editable
                              Type.Missing, // （省略可能）Notify
                              Type.Missing, // （省略可能）Converter
                              Type.Missing, // （省略可能）AddToMru
                              Type.Missing, // （省略可能）Local
                              Type.Missing  // （省略可能）CorruptLoad
                          );
                //// ワークシートを全て選択する
                var sheets = oBook.Worksheets;
                sheets.Select(Type.Missing);
                var oSheet = (Excel.Worksheet)sheets[1];
                //// A1セルを選択する
                var cells = oSheet.Cells;
                var range = ((Excel.Range)cells);
                range.Select();

                oSheet.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(cells);
                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(oBook);
                Marshal.ReleaseComObject(books);
            }
        }

        #region IDisposable support
        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// オブジェクトの後始末
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_app != null)
                    {
                        this._app.Quit();
                        Marshal.ReleaseComObject(this._app);
                        _app = null;
                    }
                    Console.WriteLine("Object disposed.");
                }
                _disposed = true;
            }
        }

        ~Stiffer()
        {
            Dispose(false);
        }
        #endregion

      

    }
}
