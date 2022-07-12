using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
//using System.Diagnostics;

namespace 席替えアプリ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //----------------------------------------------------------------------------------------------------------
        // formに枠線を引く
        private void create_line(object sender, PaintEventArgs e)
        {
            Graphics g = CreateGraphics();

            Pen blackPen = new Pen(Color.Black, 1);

            // area1
            for (int i = 0; i < 2; i++)
            {
                Point Start_point = new Point(570 + i * 90, 100);
                Point end_point = new Point(570 + i * 90, 340);

                g.DrawLine(blackPen, Start_point, end_point);
            };
            for (int i = 0; i < 5; i++)
            {
                Point Start_point = new Point(570, 100 + i * 60);
                Point end_point = new Point(660, 100 + i * 60);

                g.DrawLine(blackPen, Start_point, end_point);
            };


            // area2
            for (int i = 0; i < 3; i++)
            {
                Point Start_point = new Point(320 + i * 80, 70);
                Point end_point = new Point(320 + i * 80, 250);

                g.DrawLine(blackPen, Start_point, end_point);
            };
            for (int i = 0; i < 4; i++)
            {
                Point Start_point = new Point(320, 70 + i * 60);
                Point end_point = new Point(480, 70 + i * 60);

                g.DrawLine(blackPen, Start_point, end_point);
            };


            // area3
            for (int i = 0; i < 3; i++)
            {
                Point Start_point = new Point(320 + i * 80, 280);
                Point end_point = new Point(320 + i * 80, 460);

                g.DrawLine(blackPen, Start_point, end_point);
            };
            for (int i = 0; i < 4; i++)
            {
                Point Start_point = new Point(320, 280 + i * 60);
                Point end_point = new Point(480, 280 + i * 60);

                g.DrawLine(blackPen, Start_point, end_point);
            };


            // area4
            for (int i = 0; i < 3; i++)
            {
                Point Start_point = new Point(120 + i * 80, 70);
                Point end_point = new Point(120 + i * 80, 250);

                g.DrawLine(blackPen, Start_point, end_point);
            };

            for (int i = 0; i < 4; i++)
            {
                Point Start_point = new Point(120, 70 + i * 60);
                Point end_point = new Point(280, 70 + i * 60);

                g.DrawLine(blackPen, Start_point, end_point);
            };


            // area5
            for (int i = 0; i < 3; i++)
            {
                Point Start_point = new Point(120 + i * 80, 280);
                Point end_point = new Point(120 + i * 80, 460);

                g.DrawLine(blackPen, Start_point, end_point);
            };
            for (int i = 0; i < 4; i++)
            {
                Point Start_point = new Point(120, 280 + i * 60);
                Point end_point = new Point(280, 280 + i * 60);

                g.DrawLine(blackPen, Start_point, end_point);
            };

            blackPen.Dispose();

            g.Dispose();
        }

        // ----------------------------------------------------------------------------------------------------------
        // 起動時に「席替え.xlsx」がデスクトップにあれば読み込む(ない場合は、excel、csvの順で探す)
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {   // 席替え後のexcel
                read_excel_after_seat_change();
            }
            catch
            {
                try
                {   // excel
                    read_initial_excel();
                }
                catch
                {
                    try
                    {   // csv
                        read_initial_csv();
                    }
                    catch
                    {
                        MessageBox.Show("ファイルが見つかりませんでした");
                    }
                }
            }
        }


        //---------------------------------------------------------------------------------------------------------
        // 席替えを実行（ラベルから取得してシャッフル）
        private void change(object sender, EventArgs e)
        {
            var max_row = 0;
            string[,] members = new string[29, 2];
            for (var i = 0; i < 29; i++)
            {
                var label_name = "label" + (i + 1).ToString();
                var sl_name = "sl" + (i + 1).ToString();
                Control c = Controls[label_name];
                Control sc = Controls[sl_name];

                if (c != null && sc != null)
                {
                    members[i, 0] = ((Label)c).Text;
                    members[i, 1] = ((Label)sc).Text;
                    max_row++;
                }
            }

            // シャッフル
            do_shuffle(members, max_row);
        }

        //--------------------------------------------------------------------------------------------------------
        // excelファイルに出力
        private void output_to_file(object sender, EventArgs e)
        {
            // デスクトップパスを取得
            string desktopDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            // インスタンス生成
            var excel = new Excel.Application();
            Excel.Workbooks workbooks = excel.Workbooks;
            Excel.Workbook workbook = workbooks.Add(Type.Missing);
            var ws = (Excel.Worksheet)excel.Worksheets[1];

            try
            {
                ws.Name = "席一覧";

                // セルに枠線を引く
                Do_on_Excel(ws);

                // セルに値を入力（for文でlabel.Textを追加)
                for (int i = 0; i < 29; i++)
                {
                    var label_name = "label" + i.ToString();
                    var sl_name = "sl" + i.ToString();
                    Control c = Controls[label_name];
                    Control sc = Controls[sl_name];

                    if (c != null && sc != null)
                    {
                        if (i < 5)
                        {
                            ws.Cells[7 + 2 * i, 9].value2 = ((Label)c).Text;
                            ws.Cells[6 + 2 * i, 9].value2 = ((Label)sc).Text;
                        }
                        else if (i < 8)
                        {
                            ws.Cells[(1 + 2 * i) - 5, 7].value2 = ((Label)c).Text;
                            ws.Cells[2 * i - 5, 7].value2 = ((Label)sc).Text;
                        }
                        else if (i < 11)
                        {
                            ws.Cells[(2 * i - 2) - 8, 6].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 3) - 8, 6].value2 = ((Label)sc).Text;
                        }
                        else if (i < 14)
                        {
                            ws.Cells[(2 * i - 1) - 8, 7].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 2) - 8, 7].value2 = ((Label)sc).Text;
                        }
                        else if (i < 17)
                        {
                            ws.Cells[(2 * i - 4) - 11, 6].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 5) - 11, 6].value2 = ((Label)sc).Text;
                        }
                        else if (i < 20)
                        {
                            ws.Cells[(2 * i - 11) - 17, 4].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 12) - 17, 4].value2 = ((Label)sc).Text;
                        }
                        else if (i < 23)
                        {
                            ws.Cells[(2 * i - 14) - 20, 3].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 15) - 20, 3].value2 = ((Label)sc).Text;
                        }
                        else if (i < 26)
                        {
                            ws.Cells[(2 * i - 13) - 20, 4].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 14) - 20, 4].value2 = ((Label)sc).Text;
                        }
                        else if (i < 29)
                        {
                            ws.Cells[(2 * i - 16) - 23, 3].value2 = ((Label)c).Text;
                            ws.Cells[(2 * i - 17) - 23, 3].value2 = ((Label)sc).Text;
                        };
                    }
                };

                excel.Application.Visible = false;
                excel.Application.DisplayAlerts = false;

                // 保存
                workbook.SaveAs($@"{desktopDirectoryPath}\席替え.xlsx");

                // excelを正常に終了する
                exit(excel, workbook, workbooks, ws);

                // メッセージを出力
                MessageBox.Show("デスクトップに「席替え.xlsx」で保存しました");
            }
            catch
            {
                MessageBox.Show("ファイルを閉じてから実行してください");

                exit(excel, workbook, workbooks, ws);
            }
        }

        // ------------------------------------------------------------------------------------------------------------
        // 席替え後のファイルから読み込み
        public void read_excel_after_seat_change()
        {
            string desktopDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string file_path = $@"{desktopDirectoryPath}\席替え.xlsx";

            // インスタンス生成
            var excel = new Excel.Application();
            Excel.Workbooks workbooks = excel.Workbooks;
            Excel.Workbook workbook = workbooks.Open(file_path);
            var ws = (Excel.Worksheet)excel.Worksheets[1];

            string[,] read_excel = new string[29, 2];
            bring_from_excel(read_excel, ws);

            // ラベルに値を入れる
            for (int i = 0; i < read_excel.GetLength(0); i++)
            {
                var label_name = "label" + i.ToString();
                Control c = Controls[label_name];
                if (c != null)
                {
                    ((Label)c).Text = read_excel[i, 0];
                }

                // ラベルに学生番号を与える
                var sl_name = "sl" + i.ToString();
                Control sc = Controls[sl_name];
                if (sc != null)
                {
                    ((Label)sc).Text = read_excel[i, 1];
                }
            };

            exit(excel, workbook, workbooks, ws);
        }

        //----------------------------------------------------------------------------------------------------------
        // excelファイルから読み込み（初期配置）
        public void read_initial_excel()
        {
            // ファイルパスを取得
            string desktopDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string file_path = $@"{desktopDirectoryPath}\学生.xlsx";

            // インスタンスを生成
            var excel = new Excel.Application();
            Excel.Workbooks workbooks = excel.Workbooks;
            Excel.Workbook workbook = workbooks.Open(file_path);
            var ws = (Excel.Worksheet)excel.Worksheets[1];

            // excelの最大行を取得
            var max_row = ws.UsedRange.Rows.Count;

            //excelファイルから氏名と学籍番号を取得
            string[,] members = new string[max_row, 2];
            for (var i = 0; i < max_row; i++)
            {
                if (ws.Cells[i + 1, 1].value2 != null && ws.Cells[i + 1, 2].value2 != null)
                {
                    members[i, 0] = ws.Cells[i + 1, 1].value2;                             // 氏名

                    if ((ws.Cells[i + 1, 2].value2.ToString()).Length < 2)                 // 学籍番号
                    {
                        members[i, 1] = "B002100" + (ws.Cells[i + 1, 2].value2).ToString();
                    }
                    else
                    {
                        members[i, 1] = "B00210" + (ws.Cells[i + 1, 2].value2).ToString();
                    };
                }
                
                else
                {
                    members[i, 0] = "label" + (i + 1).ToString();
                    members[i, 1] = "sl" + (i + 1).ToString();
                }
            };

            // ラベルに学生名と学籍番号を与える
            put_values_in_the_labels(members);

            exit(excel, workbook, workbooks, ws);

        }

        // ---------------------------------------------------------------------------------------------------------
        // csvファイルから読み込み（初期配置）
        public void read_initial_csv()
        {
            string desktopDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string file_path = $@"{desktopDirectoryPath}\学生.csv";

            FileAttributes fas = File.GetAttributes(file_path);
            fas = fas & ~FileAttributes.ReadOnly;
            File.SetAttributes(file_path, fas);

            List<string[]> lists = new List<string[]>();
            var final_row = 0;

            // Shift_JISで読み込み
            var reader = new StreamReader(File.OpenRead(file_path), Encoding.GetEncoding("Shift_JIS"));

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                lists.Add(line.Split(','));
                final_row += 1;
            }

            string[,] members = new string[final_row, 2];
            for (var i = 0; i < final_row; i++)
            {
                if (lists[i][0] != "" && lists[i][1] != "")
                {
                    members[i, 0] = lists[i][0];                         // 学生名
                    if (lists[i][1].Length < 2)                          // 学籍番号
                    {
                        members[i, 1] = "B002100" + lists[i][1];
                    }
                    else
                    {
                        members[i, 1] = "B00210" + lists[i][1];
                    }
                }
                else 
                {
                    members[i, 0] = "label" + (i + 1).ToString();
                    members[i, 1] = "sl" + (i + 1).ToString();
                }
            }

            // ラベルに学生名と学籍番号を与える
            put_values_in_the_labels(members);
        }

        //----------------------------------------------------------------------------------------------------------
        // 配列のシャッフル
        public void do_shuffle(string[,] main_arr, int max_num)
        {
            // シャッフル用配列
            int[] tmp = new int[max_num];
            for (var i = 0; i < max_num; i++)
            {
                tmp[i] = i;
            };

            // tmpをシャッフル
            Random r = new Random();
            var rnd = r.Next(99, 200);
            var index = tmp.OrderBy(x => Guid.NewGuid()).ToArray();
            for (var i = 0; i < rnd; i++)
            {
                index = tmp.OrderBy(x => Guid.NewGuid()).ToArray();
            }


            // ラベルに学生名と学生番号を与える
            for (int i = 0; i < max_num + 1; i++)
            {
                var label_name = "label" + i.ToString();
                var sl_name = "sl" + i.ToString();
                Control c = Controls[label_name];
                Control sc = Controls[sl_name];

                if (c != null && sc != null)
                {
                    ((Label)c).Text = main_arr[index[i - 1], 0];
                    ((Label)sc).Text = main_arr[index[i - 1], 1];
                }
            };
        }

        //---------------------------------------------------------------------------------------------------------
        // excelでの作業
        public void Do_on_Excel(Excel.Worksheet ws)
        {
            Excel.Borders borders;
            Excel.Border border;

            // 枠線を引く
            var range_line = ws.Range[ws.Cells[8, 9], ws.Cells[15, 9]];
            for (var i = 0; i < 5; i++)
            {
                if (i == 1)
                {
                    range_line = ws.Range[ws.Cells[5, 6], ws.Cells[10, 7]];
                }
                else if (i == 2)
                {
                    range_line = ws.Range[ws.Cells[12, 6], ws.Cells[17, 7]];
                }
                else if (i == 3)
                {
                    range_line = ws.Range[ws.Cells[5, 3], ws.Cells[10, 4]];
                }
                else if (i == 4)
                {
                    range_line = ws.Range[ws.Cells[12, 3], ws.Cells[17, 4]];
                };

                borders = range_line.Borders;

                for (var j = 0; j < 4; j++)
                {
                    if (j == 0)
                    {
                        border = borders[Excel.XlBordersIndex.xlEdgeLeft];
                    }
                    else if (j == 1)
                    {
                        border = borders[Excel.XlBordersIndex.xlEdgeRight];
                    }
                    else if (j == 2)
                    {
                        border = borders[Excel.XlBordersIndex.xlEdgeTop];
                    }
                    else
                    {
                        border = borders[Excel.XlBordersIndex.xlEdgeBottom];
                    }

                    border.LineStyle = Excel.XlLineStyle.xlContinuous;

                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

                    // 文字サイズを20に設定
                    range_line.Font.Size = 20;
                };

                // セルの幅を20に設定
                range_line.Columns.ColumnWidth = 20;
            };

            // 列サイズの調整
            ws.Range[ws.Cells[5, 1], ws.Cells[17, 9]].Rows.RowHeight = 33;

            // 学籍番号のセルの調整(左揃え、下ぞろえ)
            for (var i = 0; i < 28; i++)
            {
                if (i < 4)
                {
                    ws.Cells[8 + 2 * i, 9].Font.Size = 13;
                    ws.Cells[8 + 2 * i, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[8 + 2 * i, 9].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 7)
                {
                    ws.Cells[3 + 2 * (i - 3), 7].Font.Size = 13;
                    ws.Cells[3 + 2 * (i - 3), 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[3 + 2 * (i - 3), 7].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 10)
                {
                    ws.Cells[3 + 2 * (i - 6), 6].Font.Size = 13;
                    ws.Cells[3 + 2 * (i - 6), 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[3 + 2 * (i - 6), 6].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 13)
                {
                    ws.Cells[10 + 2 * (i - 9), 7].Font.Size = 13;
                    ws.Cells[10 + 2 * (i - 9), 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[10 + 2 * (i - 9), 7].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 16)
                {
                    ws.Cells[10 + 2 * (i - 12), 6].Font.Size = 13;
                    ws.Cells[10 + 2 * (i - 12), 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[10 + 2 * (i - 12), 6].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 19)
                {
                    ws.Cells[3 + 2 * (i - 15), 4].Font.Size = 13;
                    ws.Cells[3 + 2 * (i - 15), 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[3 + 2 * (i - 15), 4].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 22)
                {
                    ws.Cells[3 + 2 * (i - 18), 3].Font.Size = 13;
                    ws.Cells[3 + 2 * (i - 18), 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[3 + 2 * (i - 18), 3].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 25)
                {
                    ws.Cells[10 + 2 * (i - 21), 4].Font.Size = 13;
                    ws.Cells[10 + 2 * (i - 21), 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[10 + 2 * (i - 21), 4].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                else if (i < 28)
                {
                    ws.Cells[10 + 2 * (i - 24), 3].Font.Size = 13;
                    ws.Cells[10 + 2 * (i - 24), 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Cells[10 + 2 * (i - 24), 3].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }

            };

            // 教卓を配置
            ws.Cells[20, 6].value2 = "教卓";
            ws.Cells[20, 6].Font.Size = 20;
            range_line = ws.Range[ws.Cells[20, 4], ws.Cells[20, 7]];
            borders = range_line.Borders;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            
            // 題名を設置
            ws.Cells[3, 4].value2 = "情報工学科・情報システム科座席表";
            ws.Cells[3, 4].Font.Size = 25;

            // フォームから日付を取得してセルに書き込む
            ws.Cells[3, 9].value2 = $"{DateTime.Now.Year}/{DateTime.Now.Month}/{DateTime.Now.Day}";
            ws.Cells[3, 9].Font.Size = 15;

        }

        //----------------------------------------------------------------------------------------------------------
        // 席の変更（席替え後）
        private void replace_seat(object sender, EventArgs e)
        {
            // フォームから取得
            string[,] members = new string[29, 2];
            for (var i = 0; i < 29; i++)
            {
                var label_name = "label" + (i + 1).ToString();
                var sl_name = "sl" + (i + 1).ToString();
                Control c = Controls[label_name];
                Control sc = Controls[sl_name];

                if (c != null && sc != null)
                {
                    members[i, 0] = ((Label)c).Text;
                    members[i, 1] = ((Label)sc).Text;
                }
            }            

            // 入れ替え用テキストボックスから学生番号を取得
            string[] num_of_moved_people = new string[8];
            var count_moved_people = 0;
            
            for (var i = 0; i < num_of_moved_people.Length; i += 2)
            {
                var textbox_name1 = "textBox" + (i + 1).ToString();
                var textbox_name2 = "textBox" + (i + 2).ToString();
                Control t1 = Controls[textbox_name1];
                Control t2 = Controls[textbox_name2];

                if (t1 != null && t2 != null)
                {
                    //　入力されている場合
                    if (((TextBox)t1).Text != "" && ((TextBox)t2).Text != "")
                    {
                        if (((TextBox)t1).Text.Length < 2)
                        {
                            num_of_moved_people[i] = "B002100" + ((TextBox)t1).Text;
                        }
                        else
                        {
                            num_of_moved_people[i] = "B00210" + ((TextBox)t1).Text;
                        }
                        if (((TextBox)t2).Text.Length < 2)
                        {
                            num_of_moved_people[i + 1] = "B002100" + ((TextBox)t2).Text;
                        }
                        else
                        {
                            num_of_moved_people[i + 1] = "B00210" + ((TextBox)t2).Text;
                        }

                        count_moved_people += 2;
                    }
                    else 
                    {
                        break;
                    }
                }
            }

            // 2次元配列を1次元配列に戻す
            string[] read_excel_1dim = members.Cast<string>().ToArray();

            // read_excelの中からnum_of_moved_peopleに該当する人を探す
            for (var i = 0; i < count_moved_people; i += 2)
            {
                // 対象のインデックスを検索して入れ替える
                var idx1 = Array.IndexOf(read_excel_1dim, num_of_moved_people[i]);
                var idx2 = Array.IndexOf(read_excel_1dim, num_of_moved_people[i + 1]);

                // 見つからなかった場合
                if (idx1 == -1 || idx2 == -1)
                {
                    MessageBox.Show("存在しない学籍番号が入力されています");
                }
                else
                {
                    idx1 /= 2;
                    idx2 /= 2;

                    // 入れ替え
                    (members[idx1, 0], members[idx2, 0]) = (members[idx2, 0], members[idx1, 0]);
                    (members[idx1, 1], members[idx2, 1]) = (members[idx2, 1], members[idx1, 1]);
                }
            };

            put_values_in_the_labels(members);
        }

        //----------------------------------------------------------------------------------------------------------
        // ファイルから読み込み（初期配置）
        private void readbutton_Click(object sender, EventArgs e)
        {
            try
            {
                // excel
                read_initial_excel();
            }
            catch
            {
                try
                {
                    // csv
                    read_initial_csv();
                }
                catch
                {
                    MessageBox.Show("ファイルが見つかりませんでした");
                }
            }
        }

        //-------------------------------------------------------------------------------------------------------------
        //学生名と学生番号をラベルに与える
        public void put_values_in_the_labels(string[,] array)
        {
            for (int i = 0; i < array.GetLength(0) + 1; i++)
            {
                var label_name = "label" + i.ToString();
                var sl_name = "sl" + i.ToString();
                Control c = Controls[label_name];
                Control sc = Controls[sl_name];

                if (c != null && sc != null)
                {
                    // 
                    if (array[i - 1, 0] == label_name || array[i - 1, 1] == sl_name)
                    {
                        ((Label)c).Text = "";
                        ((Label)sc).Text = "";
                    }
                    else
                    {
                        ((Label)c).Text = array[i - 1, 0];
                        ((Label)sc).Text = array[i - 1, 1];
                    }
                }
            };
        }

        // ------------------------------------------------------------------------------------------------------------
        // 席替え後の席配置をラベルに表示
        public void bring_from_excel(string[,] array, Excel.Worksheet ws)
        {
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if (i < 5)
                {
                    array[i, 0] = ws.Cells[7 + 2 * i, 9].value2;
                    array[i, 1] = ws.Cells[6 + 2 * i, 9].value2;
                }
                else if (i < 8)
                {
                    array[i, 0] = ws.Cells[(1 + 2 * i) - 5, 7].value2;
                    array[i, 1] = ws.Cells[2 * i - 5, 7].value2;
                }
                else if (i < 11)
                {
                    array[i, 0] = ws.Cells[(2 * i - 2) - 8, 6].value2;
                    array[i, 1] = ws.Cells[(2 * i - 3) - 8, 6].value2;
                }
                else if (i < 14)
                {
                    array[i, 0] = ws.Cells[(2 * i - 1) - 8, 7].value2;
                    array[i, 1] = ws.Cells[(2 * i - 2) - 8, 7].value2;
                }
                else if (i < 17)
                {
                    array[i, 0] = ws.Cells[(2 * i - 4) - 11, 6].value2;
                    array[i, 1] = ws.Cells[(2 * i - 5) - 11, 6].value2;
                }
                else if (i < 20)
                {
                    array[i, 0] = ws.Cells[(2 * i - 11) - 17, 4].value2;
                    array[i, 1] = ws.Cells[(2 * i - 12) - 17, 4].value2;
                }
                else if (i < 23)
                {
                    array[i, 0] = ws.Cells[(2 * i - 14) - 20, 3].value2;
                    array[i, 1] = ws.Cells[(2 * i - 15) - 20, 3].value2;
                }
                else if (i < 26)
                {
                    array[i, 0] = ws.Cells[(2 * i - 13) - 20, 4].value2;
                    array[i, 1] = ws.Cells[(2 * i - 14) - 20, 4].value2;
                }
                else if (i < 29)
                {
                    array[i, 0] = ws.Cells[(2 * i - 16) - 23, 3].value2;
                    array[i, 1] = ws.Cells[(2 * i - 17) - 23, 3].value2;
                };
            };

        }

        //------------------------------------------------------------------------------------------------------------
        // excelの正常終了
        public void exit(Excel.Application app, Excel.Workbook workbook, Excel.Workbooks workbooks, Excel.Worksheet worksheet)
        {
            // appの終了の前に破棄する
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);

            // appの終了前にガベージコレクションを強制
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // appを終了
            app.Quit();

            // appオブジェクトを破棄
            Marshal.ReleaseComObject(app);

            // appオブジェクトのガベージコレクションを強制
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        //----------------------------------------------------------------------------------------------------------
        // アプリケーションの終了
        private void app_exit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("終了しますか？", "確認",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
