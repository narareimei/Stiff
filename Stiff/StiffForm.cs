using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Stiff
{
    public partial class StiffForm : Form
    {
        private Stiffer         stiffer;

        private DataTable       excelFiles;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public StiffForm()
        {
            InitializeComponent();
        }

        #region イベントハンドラ

        /// <summary>
        /// フォーム初期化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StiffForm_Load(object sender, EventArgs e)
        {
            // 
            stiffer = Stiffer.GetInstance();

            // グリッド初期化
            {
                bookGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                bookGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            }
            {
                var dt = new DataTable();

                // カラム定義
                dt.Columns.Add(new DataColumn("Seq",        typeof(int)));
                dt.Columns.Add(new DataColumn("Path",       typeof(string)));
                dt.Columns.Add(new DataColumn("File",       typeof(string)));
                dt.Columns.Add(new DataColumn("Author",     typeof(string)));   // 作成者
                dt.Columns.Add(new DataColumn("Title",      typeof(string)));   // タイトル
                dt.Columns.Add(new DataColumn("Subject",    typeof(string)));   // サブジェクト
                dt.Columns.Add(new DataColumn("Update",     typeof(string)));   // 更新日時
                dt.Columns.Add(new DataColumn("Company",    typeof(string)));   // 会社
                dt.Columns.Add(new DataColumn("Manager",    typeof(string)));   // 管理者
                // プライマリキー設定
                var pk = new DataColumn[1];
                pk[0] = dt.Columns["File"];
                dt.PrimaryKey = pk;
                this.excelFiles = dt;
            }
            bookGrid.DataSource = this.excelFiles;
        }

        #region ドラッグアンドドロップ

        /// <summary>
        /// ドラッグアンドドロップ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StiffForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {

                // ドラッグ中のファイルやディレクトリの取得
                string[] drags = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach (string d in drags)
                {
                    if (!System.IO.File.Exists(d))
                    {
                        // ファイル以外であればイベント・ハンドラを抜ける
                        return;
                    }
                }
                e.Effect = DragDropEffects.Copy;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StiffForm_DragDrop(object sender, DragEventArgs e)
        {
            // ドラッグ＆ドロップされたファイル
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            var list = getBookInformations(files);
            foreach (var info in list)
            {
                addBookInfo(info);
            }
            return;
        }

        #endregion //ドラッグアンドドロップ

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            var cd = System.IO.Directory.GetCurrentDirectory();

            // ブック情報をデータテーブルへ格納してみる
            var info = stiffer.GetInformations(cd + @"\TestBook.xlsx");
            {
                var row = this.excelFiles.NewRow();

                row["Seq"       ] = excelFiles.Rows.Count + 1;
                row["Path"      ] = "";
                row["File"      ] = info.FileName;
                row["Author"    ] = info.Author;
                row["Title"     ] = info.Title;
                row["Subject"   ] = info.Subject;
                row["Update"    ] = info.LastSaveTime;
                row["Company"   ] = info.Company;
                row["Manager"   ] = info.Manager;
                excelFiles.Rows.Add(row);
            }
        }

        #endregion //イベントハンドラ

        #region ロジック

        /// <summary>
        /// ブック情報をデータテーブルへ追加する
        /// </summary>
        /// <param name="info"></param>
        private void addBookInfo(BookInfo info)
        {
            var row = this.excelFiles.NewRow();

            row["Seq"       ] = excelFiles.Rows.Count + 1;
            row["Path"      ] = "";
            row["File"      ] = info.FileName;
            row["Author"    ] = info.Author;
            row["Title"     ] = info.Title;
            row["Subject"   ] = info.Subject;
            row["Update"    ] = info.LastSaveTime;
            row["Company"   ] = info.Company;
            row["Manager"   ] = info.Manager;
            excelFiles.Rows.Add(row);
            return;
        }

        /// <summary>
        /// ブック情報の取得
        /// </summary>
        /// <param name="files">ファイル名配列</param>
        private BookInfo []  getBookInformations(string[] files)
        {
            var list = new List<BookInfo>();

            foreach( var file in files ) 
            {
                var info = this.stiffer.GetInformations(file);
                list.Add(info);
            }
            return list.ToArray();
        }

        #endregion //ロジック
    }
}
