using System;
using System.Drawing;
using System.Windows.Forms;

namespace DragAndDrop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //ListBox1のDragEnterイベントハンドラ
        private void ListBox1_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            //コントロール内にドラッグされたとき実行される
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                //ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
                e.Effect = DragDropEffects.Copy;
            else
                //ファイル以外は受け付けない
                e.Effect = DragDropEffects.None;
        }

        //ListBox1のDragDropイベントハンドラ
        private void ListBox1_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            //コントロール内にドロップされたとき実行される
            //ドロップされたすべてのファイル名を取得する
            string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            //ListBoxに追加する
            listBox1.Items.AddRange(fileName);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                DialogResult dr = MessageBox.Show("指定したファイルを削除しますか？", "【ファイル削除】", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                if (dr == DialogResult.OK)
                {
                    System.IO.File.Delete(listBox1.SelectedItem.ToString());
                    listBox1.Items.Remove(listBox1.SelectedItem);
                }
            }
            else
            {
                MessageBox.Show("選択されたアイテムがありません。", "【ファイル削除】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            // 選択したファイルが存在する場合
            if (listBox1.SelectedItem != null)
            {
                // 拡張子を取得
                string extensionKey = System.IO.Path.GetExtension(listBox1.SelectedItem.ToString()).ToLower();

                switch (extensionKey)
                {
                    case ".pdf":     // PDFファイル
                        // ブラウザコントロールの作成
                        axAcroPDF1.Visible = true;
                        printPreviewControl1.Visible = false;

                        axAcroPDF1.LoadFile(listBox1.SelectedItem.ToString());
                        break;

                    case ".csv":     // CSVファイル
                        axAcroPDF1.Visible = false;
                        printPreviewControl1.Visible = true;

                        printPreviewControl1.Document = printDocument1;

                        break;

                    case ".xlsx":    // EXCELファイル
                        axAcroPDF1.Visible = false;
                        printPreviewControl1.Visible = true;

                        break;

                    default:
                        axAcroPDF1.Visible = false;
                        printPreviewControl1.Visible = false;

                        /* 特定できてない拡張子が来た場合を想定 */
                        break;
                }
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //画像を読み込む
            Image img = Image.FromFile(listBox1.SelectedItem.ToString());
            //画像を描画する
            e.Graphics.DrawImage(img, 0, 0, img.Width, img.Height);
            //次のページがないことを通知する
            e.HasMorePages = false;
            //後始末をする
            img.Dispose();
        }
    }
}
