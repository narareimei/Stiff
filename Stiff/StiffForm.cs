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


        public StiffForm()
        {
            InitializeComponent();
        }

        private void StiffForm_Load(object sender, EventArgs e)
        {
            // 
            stiffer = Stiffer.GetInstance();

            // グリッド初期化
            {
                var dt = new DataTable();

                // カラム定義
                dt.Columns.Add(new DataColumn( "Seq",       typeof(int)));
                dt.Columns.Add(new DataColumn( "Path",      typeof(string)));
                dt.Columns.Add(new DataColumn( "File",      typeof(string)));
                dt.Columns.Add(new DataColumn( "Author",    typeof(string)));   // 作成者
                dt.Columns.Add(new DataColumn( "Title",     typeof(string)));   // タイトル
                dt.Columns.Add(new DataColumn( "Subject",   typeof(string)));   // サブジェクト
                dt.Columns.Add(new DataColumn( "Update",    typeof(DateTime))); // 更新日時
                dt.Columns.Add(new DataColumn( "Company",   typeof(string)));   // 会社
                dt.Columns.Add(new DataColumn( "Manager",   typeof(string)));   // 管理者
                this.excelFiles = dt;
            }
            bookGrid.DataSource = this.excelFiles;
        }

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

        private void StiffForm_DragDrop(object sender, DragEventArgs e)
        {
            // ドラッグ＆ドロップされたファイル
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);


        }

        private void button1_Click(object sender, EventArgs e)
        {
            var cd = System.IO.Directory.GetCurrentDirectory();

            // ブック情報をデータテーブルへ格納してみる
            var info = stiffer.GetInformations(cd + @"\TestBook.xlsx");
            {
                var row = this.excelFiles.NewRow();

                row["Seq"       ] = 1;
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


    }
}
