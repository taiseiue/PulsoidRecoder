using CefSharp;
using CefSharp.OffScreen;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PulsoidRecoder
{
    class Program
    {
        static bool preseCancel= false;
        static void Main(string[] args)
        {
            Console.WriteLine("PulsoidRecoder (c)WSOFT 2020");
            Console.CancelKeyPress += Console_CancelKeyPress;
            CefSettings settings = new CefSettings();
            settings.Locale = "ja";
            settings.AcceptLanguageList = "ja-JP";
            
            Cef.Initialize(settings, performDependencyCheck: false, browserProcessHandler: null);

            bool breakflag = false;

            while (!breakflag)
            {
                Console.WriteLine("PulsoidRecoder メインメニュー");
                Console.WriteLine("実行したい操作を選択してください");
                Console.Write("[R:記録モード / E:エクセル出力モード / V:閲覧モード / QまたはEsc:終了]");
                switch (Console.ReadKey().Key)
                {
                    case ConsoleKey.Escape:
                        {
                           
                            breakflag = true;
                            break;
                        }
                    case ConsoleKey.Q:
                        {
                          
                            breakflag = true;
                            break;
                        }
                    case ConsoleKey.E:
                        {
                            Console.WriteLine("読み込むファイル名を入力してください");
                            Console.Write("ファイルパス>");
                            string filepath = Console.ReadLine();
                            Console.Write("解析しています.");
                            BPMRecordCollection collection = LoadFile(filepath);
                            XLWorkbook workbook = new XLWorkbook();
                            IXLWorksheet worksheet = workbook.AddWorksheet();
                            worksheet.Name = "Pulseroid RecordLog";
                            IXLCell c = worksheet.Cell(1, 1);
                            c.Value = "日時";
                            IXLCell c1 = worksheet.Cell(1, 2);
                            c.Value = "心拍数";
                            IXLCell c2 = worksheet.Cell(1,3);
                            c.Value = "最高心拍数";
                            IXLCell c3 = worksheet.Cell(1, 4);
                            c.Value = "平均心拍数";
                            IXLCell c4 = worksheet.Cell(1, 5);
                            c.Value = "最低心拍数";
                            IXLCell c5 = worksheet.Cell(2, 3);
                            c.Value = collection.Maxbpm;
                            IXLCell c6 = worksheet.Cell(2, 4);
                            c.Value = collection.Avgbpm;
                            IXLCell c7 = worksheet.Cell(2, 5);
                            c.Value = collection.Minbpm;
                            int cx = 2;
                            int count = 0;
                            foreach (BPMRecord bpm in collection.Records)
                            {
                               
                                 
                                    IXLCell x = worksheet.Cell(cx, 1);
                                x.Value = bpm.Month + "/" + bpm.Day + " " + bpm.Hour + ":" + bpm.Minute + ":" + bpm.Second;
                                    IXLCell x2 = worksheet.Cell(cx, 2);
                                x2.Value = bpm.BPM;
                                count++;
                                if (count == 60)
                                {
                                    Console.Write(".");
                                    count = 0;
                                }
                                cx++;


                            }
                            Console.WriteLine();
                            Console.WriteLine("解析完了");
                            Console.WriteLine("保存先を入力してください");
                            Console.Write("ファイルパス(.xlsx)>");
                            string savefilepath = Console.ReadLine();
                            if (Path.GetExtension(savefilepath) != ".xlsx")
                            {
                                savefilepath += ".xlsx";
                            
                            }
                            Console.WriteLine("保存しています...");
                            workbook.SaveAs(savefilepath);
                            Console.WriteLine("保存完了");
                            break;
                        }
                    case ConsoleKey.V:
                        {
                            Console.WriteLine("読み込むファイル名を入力してください");
                            Console.Write("ファイルパス>");
                            string filepath = Console.ReadLine();
                            BPMRecordCollection collection = LoadFile(filepath);
                            Console.WriteLine("記録:"+collection.Records.Length+"個 最高心拍数:"+collection.Maxbpm+"BPM 平均心拍数:"+collection.Avgbpm+"BPM 最低心拍数:"+collection.Minbpm+"BPM");
                            Console.WriteLine();
                            foreach(BPMRecord bpm in collection.Records)
                            {
                                Console.WriteLine(bpm.Month+"/"+bpm.Day+" "+bpm.Hour+":"+bpm.Minute+":"+bpm.Second+"  "+bpm.BPM+"BPM");
                            }
                            break;
                        }
                    case ConsoleKey.R:
                        {
                            Console.WriteLine();
                            string WidetUrl = "";
                              Console.WriteLine("PulsoidのウィジェットUrlを入力してください");
                                Console.Write("URL>");
                                WidetUrl = Console.ReadLine();
                            int span = 1;
                            Console.WriteLine("何秒おきに記録するか設定してください");
                            Console.Write("時間(秒)>");
                            int.TryParse(Console.ReadLine(),out span);
                            if (span < 1)
                            {
                                span = 1;
                            }
                             
                            string cachePath = Path.GetFullPath("cache");
                            var browserSettings = new BrowserSettings();
                            //毎秒一枚更新されればOK
                            browserSettings.WindowlessFrameRate = 1;
                            var requestContextSettings = new RequestContextSettings { CachePath = cachePath };
                            
                           
                            using (var requestContext = new RequestContext(requestContextSettings))
                            using (var browser = new ChromiumWebBrowser(WidetUrl, browserSettings, requestContext))
                            {
                                List<BPMRecord> records = new List<BPMRecord>();

                                //起動中は待機
                                while (!browser.IsBrowserInitialized) { }
                                //読み込み中は待機
                                while (browser.IsLoading) { }
                                //BPMが表示されていない間は待機
                                while (GetBPM(browser)==0) { }

                                int maxbpm = 0;
                                int minbpm = 999;
                                double bpmcount = 0;

                              int cct=  Console.CursorTop;
                                int ccl = Console.CursorLeft;
                                Console.WriteLine("[***BPM  0秒]");

                                    //記録スタート
                                    while (true)
                                    {
                                        int bpm = GetBPM(browser);
                                    if (maxbpm < bpm) { maxbpm = bpm; }
                                    if (minbpm > bpm) { minbpm = bpm; }
                                    bpmcount += bpm;
                                        records.Add(new BPMRecord(DateTime.Now, bpm));
                                    Task.Run(() =>
                                    {
                                        Console.SetCursorPosition(ccl, cct);
                                        Console.WriteLine("[ " + bpm + "BPM  最大:" + maxbpm + "BPM 平均:"+ (int)(bpmcount / records.Count) + "BPM 最小:" + minbpm + "BPM " + records.Count + "個の記録 ]");
                                    });
                                        if (preseCancel) { break; }
                                        Thread.Sleep(span*1000);
                                    }
                                Console.WriteLine("記録を終了しました");
                                Console.WriteLine("記録の保存先を入力してください");
                                Console.Write("ファイルパス>");
                                string savepath = Console.ReadLine();
                                Console.WriteLine("保存しています...");
                               
                                SaveFile(new BPMRecordCollection(records.ToArray(),maxbpm,minbpm,(int)(bpmcount/records.Count)),savepath);
                               
                                Console.WriteLine("保存完了");
                            }
                            break;

                        }
                }
            }
        }
       
       static void SaveFile(BPMRecordCollection recods, string path)
        {
            FileStream fs = new FileStream(path,
                FileMode.Create,
                FileAccess.Write);
            BinaryFormatter bf = new BinaryFormatter();
           
            bf.Serialize(fs, recods);
            fs.Close();
        }
        static BPMRecordCollection LoadFile(string path)
        {
            FileStream fs = new FileStream(path,
                FileMode.Open,
                FileAccess.Read);
            BinaryFormatter f = new BinaryFormatter();
           
            object obj = f.Deserialize(fs);
            fs.Close();
            if(obj is BPMRecordCollection recods)
            {
                return recods;
            }
            return null;
        }
        private static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            e.Cancel = true;
            preseCancel = true;
        }

        static int GetBPM(ChromiumWebBrowser browser)
        {
            Task<string> task = browser.GetTextAsync();
            task.Wait();
            int re = 0;
            int.TryParse(task.Result,out re);
            return re;
        }
    }
}
