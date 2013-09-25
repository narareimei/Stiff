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
                dt.Columns.Add(new DataColumn( "seq",       typeof(int)));
                dt.Columns.Add(new DataColumn( "path",      typeof(string)));
                dt.Columns.Add(new DataColumn( "file",      typeof(string)));
                dt.Columns.Add(new DataColumn( "author",    typeof(string)));   // 作成者
                dt.Columns.Add(new DataColumn( "title",     typeof(string)));   // タイトル
                dt.Columns.Add(new DataColumn( "subtitle",  typeof(string)));   // サブタイトル
                dt.Columns.Add(new DataColumn( "update",    typeof(DateTime))); // 更新日時
                dt.Columns.Add(new DataColumn( "company",   typeof(string)));   // 会社
                dt.Columns.Add(new DataColumn( "manager",   typeof(string)));   // 管理者
            }
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
            stiffer.GetInformations(@"C:\Users\Administrator\Dropbox\private\dotNet\Stiff\Stiff\TestBook.xlsx");
        }


    }
}
