using System;
using System.Data;
using System.IO;
using System.Text;
using QuickType;
using System.Reflection;
using System.Text.RegularExpressions;

namespace VbsWriter
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            // jsonTextファイルを開く
            using (var sr = new StreamReader(@"../../file/test_template_json.txt", Encoding.GetEncoding("shift_jis")))
            {
                string jsonString = sr.ReadToEnd();
                var welcomes = Welcome.FromJson(jsonString);

                DataSet dataSet = new DataSet();
                DataTable table = new DataTable("Table");

                // プロパティの属性を取得
                PropertyInfo[] infoArrayTOAddColumns = welcomes[0].GetType().GetProperties();

                // プロパティ情報出力をループで回す
                foreach (PropertyInfo info in infoArrayTOAddColumns)
                {
                    // カラム名の追加
                    table.Columns.Add(info.Name);
                }

                // DataSetにDataTableを追加
                dataSet.Tables.Add(table);

                foreach (var welcome in welcomes)
                {
                    DataRow dr = table.NewRow();

                    PropertyInfo[] infoArray = welcome.GetType().GetProperties();
                    // プロパティ情報出力をループで回す
                    foreach (PropertyInfo info in infoArray)
                    {
                        Console.WriteLine(info.Name + ": " + info.GetValue(welcome, null));
                        dr[info.Name] = info.GetValue(welcome, null);
                    }

                    table.Rows.Add(dr);
                }

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (!Regex.IsMatch(table.Rows[i]["startTime"].ToString(), "\\d{4}"))
                    {
                        Console.WriteLine("startTimeが数字じゃない");
                        continue;
                    }

                    // ファイルから一行読み込む
                    var line = table.Rows[i];

                    // vbsファイルのタイトルを作成
                    string vbsTitle = createVbsTitleFromTableRow(line);

                    // vbsファイルの保存先パスを作成
                    string vbsPath = GetFullPathWithCurrentDirectoryAndTitleAsVBSFileName(vbsTitle);

                    // vbsファイルを作成
                    // 第２引数：毎度新しいファイルを作成するため、Appendフラグ（2番目の引数）をfalseに設定
                    // 第３引数：VBSで動作するように文字コードをshift-jisに指定
                    // ＣＳＶの複数行を書き込めるようにするため、あえてusingステートメントを使用しない。
                    using (var writer = new StreamWriter(vbsPath, false, Encoding.GetEncoding("shift_jis")))
                    {
                        // vbsの中身を作成
                        var procedure = line["procedure"].ToString().Replace("\"\" & vbCrLf & \"\"", "\" & vbCrLf & \"");
                        var vbsContent = $"val = MsgBox({procedure} ,vbSystemModal + vbExclamation , \"{vbsTitle}\")";
                        
                        // 書き込む
                        writer.WriteLine(vbsContent);

                        var nextRows = table.Select("startTime LIKE 'C" + line["startRow"] + "%' " +
                            "OR startTime LIKE 'B" + line["startRow"] + "%'");

                        Console.WriteLine(nextRows.Length);
                        while (nextRows.Length > 0)
                        {
                            foreach (var nextRow in nextRows)
                            {
                                // vbsファイルのタイトルを作成
                                string nextVbsTitle = createVbsTitleFromTableRow(nextRow);

                                // content
                                var nextProcedure = nextRow["procedure"].ToString().Replace("\"\" & vbCrLf & \"\"", "\" & vbCrLf & \"");
                                var nextVbsContent = $"If val = vbOK  Then val = MsgBox({nextProcedure} ,vbSystemModal + vbExclamation , \"{nextVbsTitle}\")";

                                // 書き込む
                                writer.WriteLine(nextVbsContent);

                            }

                            nextRows = table.Select("startTime LIKE 'C" + nextRows[0]["startRow"] + "%' " +
                            "OR startTime LIKE 'B" + nextRows[0]["startRow"] + "%'");
                        }
                    };
                }
            }
            Console.ReadLine();
        }

        private static string createVbsTitleFromTableRow(DataRow row)
        {
            // タイトル文字列を作成
            string vbsTitle = row["startTime"].ToString().Replace(":", "")
                + "_" + row["targetEnvironment"]
                + "_" + row["workTitle"].ToString().Replace(" & vbCrLf & ", "")
                .Replace("\"", "").Replace(":", "");

            // タイトルが５０文字より長い場合は先頭から５０文字まで切り取る
            return vbsTitle.Length > 50 ? vbsTitle.Substring(0, 50) : vbsTitle;
        }

        private static string GetFullPathWithCurrentDirectoryAndTitleAsVBSFileName(string vbsTitle)
        {
            // カレントディレクトリのパスを取得する
            string CurrentDir = Directory.GetCurrentDirectory();

            string fName = vbsTitle + ".vbs";

            // 保存先のVbsファイルのパスを組み立てる
            return Path.Combine(CurrentDir, fName);
        }
    }
}
