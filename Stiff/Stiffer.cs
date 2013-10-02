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
        public String Subject { get; set; }
        public String Manager { get; set; }
        public String Company { get; set; }

        public String FileName { get; set; }
        public String LastSaveTime { get; set; }
        public bool? Result { get; set; }
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
        /// Excelブックの各種情報取得
        /// </summary>
        public BookInfo GetInformations(string filename)
        {
            Excel.Workbook oBook = null;
            BookInfo info = null ;
            try
            {
                // アプリケーション起動
                this.CreateApplication();

                // ファイルオープン
                oBook = this.OpenBook(filename);
                if (oBook == null)
                {
                    return null;
                }

                // ブック情報取得および格納
                info = new BookInfo();
                info.FileName       = filename;
                info.Author         = this.GetBuiltinProperty(oBook, "Author");
                info.Title          = this.GetBuiltinProperty(oBook, "Title");
                info.Subject        = this.GetBuiltinProperty(oBook, "Subject");
                info.Manager        = this.GetBuiltinProperty(oBook, "Manager");
                info.Company        = this.GetBuiltinProperty(oBook, "Company");
                info.LastSaveTime   = this.GetBuiltinProperty(oBook, "Last Save Time");
            }
            finally
            {
                if (oBook != null)
                {
                    oBook.Close(false, filename, Type.Missing);
                    Marshal.ReleaseComObject(oBook);
                }
                oBook = null;
            }
            return info;
        }

        /// <summary>
        /// 新しい状態、プロパティで保存する
        /// </summary>
        public void Unification(BookInfo [] infomations)
        {
            // アプリケーション起動
            this.CreateApplication();

            foreach( var info in infomations)
            {
                info.Result = false;
                //
                Excel.Workbook oBook = null;
                Excel.Sheets oSheets = null;
                Excel.Worksheet oTopSheet = null;
                int topIdx = 0;

                string filename = info.FileName;
                try
                {
                    // ファイルオープン
                    oBook = this.OpenBook(filename);
                    if (oBook == null)
                    {
                        return;
                    }

                    // プロパティ設定
                    {
                        info.Title   = "";
                        info.Subject = "";
                        info.Company = "";
                        info.Manager = "";
                    }
                    this.SetInformations(oBook, info);

                    // 状態変更
                    oSheets = oBook.Worksheets;
                    for (int i = 1; i <= oSheets.Count; ++i)
                    {
                        Excel.Worksheet oSheet = null;
                        Excel.Range oCells = null;
                        Excel.Range oRange = null;

                        try
                        {
                            oSheet = (Excel.Worksheet)oSheets[i];
                            if (oSheet.Visible != Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
                            {
                                continue;
                            }
                            oSheet.Activate();

                            // 最初に表示シートを保持しておく
                            if (topIdx == 0)
                            {
                                topIdx = i;
                            }

                            // A1セル
                            oCells = oSheet.Cells;
                            oRange = ((Excel.Range)oCells[1, 1]);
                            oRange.Select();

                            // 表示倍率
                            this._app.ActiveWindow.Zoom = 100;

                            // 枠線
                            this._app.ActiveWindow.DisplayGridlines = false;

                            // 表示モード
                            this._app.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView;

                        }
                        finally
                        {
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
                        }
                    }
                    if (topIdx != 0)
                    {
                        oTopSheet = (Excel.Worksheet)oSheets[topIdx];
                        oTopSheet.Activate();
                    }

                    // 保存
                    oBook.Close(true, filename, Type.Missing);
                    Marshal.ReleaseComObject(oBook);
                    oBook = null;
                    info.Result = true;
                }
                finally
                {
                   if (oTopSheet != null)
                    {
                        Marshal.ReleaseComObject(oTopSheet);
                    }
                    oTopSheet = null;

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
                }
            }
        }

        #region private methods

        /// <summary>
        /// 
        /// </summary>
        private void CreateApplication()
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
        private Excel.Workbook OpenBook(string filename)
        {
            var oBooks = this._app.Workbooks;
            Excel.Workbook oBook = null;

            try
            {
                oBook = oBooks.Open(
                              filename,  // オープンするExcelファイル名
                              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                              Type.Missing, // （省略可能）ReadOnly (True / False )
                              Type.Missing, // （省略可能）Format
                                            //      1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                                            //      5:なし / 6:引数 Delimiterで指定された文字
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
        private string GetBuiltinProperty(Excel.Workbook oBook, string propertyName)
        {
            string strValue = "";

            try
            {
                object ps = oBook.BuiltinDocumentProperties;
                Type typeDocBuiltInProps = ps.GetType();

                //Get the property and display it.
                object oDocProp = typeDocBuiltInProps.InvokeMember("Item",
                                           BindingFlags.Default | BindingFlags.GetProperty,
                                           null, ps, new object[] { propertyName });

                Debug.Assert(oDocProp != null);
                Type typeDocAuthorProp = oDocProp.GetType();
                strValue = typeDocAuthorProp.InvokeMember("Value",
                                           BindingFlags.Default | BindingFlags.GetProperty,
                                           null, oDocProp, new object[] { }).ToString();

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
        /// Excelファイルのプロパティを設定する
        /// </summary>
        /// <param name="oBook"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private void SetBuiltinProperty(Excel.Workbook oBook, string propertyName, string value)
        {
            try
            {
                object ps = oBook.BuiltinDocumentProperties;
                Type typeDocBuiltInProps = ps.GetType();

                //Get the property and display it.
                object oDocProp = typeDocBuiltInProps.InvokeMember("Item",
                                           BindingFlags.Default | BindingFlags.GetProperty,
                                           null, ps, new object[] { propertyName });

                Debug.Assert(oDocProp != null);
                Type typeDocAuthorProp = oDocProp.GetType();
                typeDocAuthorProp.InvokeMember("Value",
                                           BindingFlags.Default | BindingFlags.SetProperty,
                                           null, oDocProp, new object[] { value });
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("The application failed with the following exception: '{0}' Stacktrace:'{1}'", ex.Message, ex.StackTrace));
            }
            finally
            {
            }
            return;
        }

        /// <summary>
        /// ブックに新しいプロパティを設定する
        /// </summary>
        /// <param name="oBook"></param>
        /// <param name="info"></param>
        private void SetInformations(Excel.Workbook oBook, BookInfo info)
        {
            this.SetBuiltinProperty(oBook, "Author" , info.Author);
            this.SetBuiltinProperty(oBook, "Title"  , info.Title);
            this.SetBuiltinProperty(oBook, "Subject", info.Subject);
            this.SetBuiltinProperty(oBook, "Manager", info.Manager);
            this.SetBuiltinProperty(oBook, "Company", info.Company);
        }

        #endregion

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
