using System;
using System.IO;

namespace VbsWriter
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            // vbsファイルの保存パスを生成
            string vbsPath = GetFullPathWithCurrentDirectoryAndTitleAsCSVFileName("test");

            // TODO:CSVファイルを変数CSVへ読み込む

            // TODO:変数CSVから、開始時間が固定された作業（固定作業）のみ抽出して変数koteiLinesに設定

            // TODO:変数koteiLinesの末尾まで以下のループ処理する。変数lineを用いる。

            // TODO:1.変数vbsContentに変数line分のMsgBox生成スクリプトを追加
            // TODO:2.変数lineから、続き作業があるか判別
            // TODO:2-1.続き作業の呼出ある場合は、変数lineに続き作業を設定
            // TODO:2-1-1.変数vbsContentにOKボタンクリック後の続きスクリプトを追加
            // TODO:2-1-2.変数vbsContentに変数line分のMsgBox生成スクリプトを追加
            // TODO:2-1-3.変数lineから、続き作業があるか判別(以下、続きがなくなるまでループ)
            // TODO:2-2.続き作業の呼出がない場合は、ファイルを生成して次のlineへ進む

            // vbsファイルの中身を作成
            var body = "ここに本文";
            var title = "ここにタイトル";
            var vbsContent = "MsgBox \"" + body + "\", vbSystemModal + vbExclamation, \"" + title + "\"";
            Console.WriteLine(vbsContent);

            // vbsファイルを作成
            using (var writer = new StreamWriter (vbsPath)){
                writer.WriteLine(vbsContent);
            }
        }

        private static string GetFullPathWithCurrentDirectoryAndTitleAsCSVFileName(string vbsTitle)
        {
            // カレントディレクトリのパスを取得する
            string CurrentDir = Directory.GetCurrentDirectory();

            string fName = vbsTitle + ".vbs";

            // 保存先のVbsファイルのパスを組み立てる
            return Path.Combine(CurrentDir, fName);
        }
    }
}
