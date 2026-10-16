using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ScreenBank {
    public static class SelfTest {
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr handle,int command);
        static int passed,failed;static StringBuilder report=new StringBuilder();
        static void Check(bool ok,string name){if(ok)passed++;else failed++;report.AppendLine((ok?"PASS ":"FAIL ")+name);}
        static IEnumerable<Control> Controls(Control root){foreach(Control child in root.Controls){yield return child;foreach(var next in Controls(child))yield return next;}}
        public static int Run(string[] args){
            string output=args.Length>0?Path.GetFullPath(args[0]):AppDomain.CurrentDomain.BaseDirectory;Directory.CreateDirectory(output);
            try{
                string sample=Path.Combine(output,"test-template.xlsx");if(File.Exists(sample))File.Delete(sample);BankFile.Template(sample);
                var sheets=BankFile.Read(sample);var map=BankFile.Guess(sheets[0].Rows[0]);var qs=BankFile.Convert(sheets[0],0,map,sample);var matcher=new Matcher(qs);
                Check(qs.Count==4 && qs[0].Options[0]=="北京" && qs[0].Answer=="A","XLSX inline strings and header mapping");
                Check(BankFile.Guess(new []{"题干","答案","选项G"})[9]==2,"G option column");
                Check(matcher.Search("1. 中 国 的 首 都 是 哪 里？\nA.北京\nB.上海\nC.广州\nD.深圳",5)[0].Question==qs[0],"OCR whitespace and question number");
                Check(matcher.Search("下列哪个数字不是质数？",5)[0].Question==qs[2],"Negation distinguishes similar questions");
                Check(matcher.Search("下列哪个数字是质数？",5)[0].Question==qs[1],"Positive question distinguishes similar questions");
                Check(matcher.Search("这是一条完全无关的新闻消息",5).Count==0,"Unrelated query has no candidates");
                Check(matcher.Search("北京",5).Count==0,"Very short OCR is rejected");
                Check(matcher.Search("中国的首都是哪座城市？",5)[0].Question==qs[0],"Fuzzy wording changes");
                Check(matcher.Search("哪座城市是中国首都？",5)[0].Question==qs[0],"Fuzzy reordered wording");
                Check(matcher.Search("以下哪一项数字并非质数？",5)[0].Question==qs[2],"Fuzzy wording keeps negation");
                Check(matcher.Search("水分子的化学式为？",5)[0].Question==qs[3],"Fuzzy additions and omissions");
                var broken=new Question{Text="测量线路电压数值",Options=new string[7]};var brokenMatcher=new Matcher(new List<Question>{broken});var brokenHits=brokenMatcher.Search("测线电值",5);
                Check(brokenHits.Count>0&&brokenHits[0].Score>=45,"Character fallback recovers broken adjacent pairs");
                Question numberA=new Question{Text="额定频率50Hz时电流如何计算",Options=new string[7]},numberB=new Question{Text="额定频率60Hz时电流如何计算",Options=new string[7]};var numericMatcher=new Matcher(new List<Question>{numberA,numberB});
                Check(numericMatcher.Search("额定频率60Hz时电流怎样计算",5)[0].Question==numberB,"Numbers distinguish fuzzy technical questions");
                var csv=BankFile.Csv("题目,答案\r\n\"第一行\n第二行,包含逗号\",\"A\"\"B\"\r\n",',');
                Check(csv.Count==2 && csv[1][0].Contains("\n") && csv[1][1]=="A\"B","CSV quoted multiline and escaped quote");
                bool malformed=false;try{BankFile.Csv("\"x",',');}catch{malformed=true;}Check(malformed,"Malformed CSV is rejected");
                string real=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"电气专业试题库（第六周）.xlsx");
                var realSheets=BankFile.Read(real);var realMap=BankFile.Guess(realSheets[0].Rows[0]);var realQs=BankFile.Convert(realSheets[0],0,realMap,real);var realMatcher=new Matcher(realQs);
                Check(realQs.Count==200 && realMap[0]==2 && realMap[1]==10,"User XLSX shared strings, 200 questions and actual columns");
                int exact=0,noisy=0;foreach(var q in realQs){var hits=realMatcher.Search(q.Text,5);if(hits.Count>0 && Matcher.Normalize(hits[0].Question.Text)==Matcher.Normalize(q.Text))exact++;
                    string ocrText="12. "+string.Join(" ",q.Text.Select(c=>c.ToString()))+"\n"+string.Join("\n",q.Options.Select((o,i)=>((char)('A'+i))+". "+o));hits=realMatcher.Search(ocrText,5);if(hits.Count>0&&Matcher.Normalize(hits[0].Question.Text)==Matcher.Normalize(q.Text))noisy++;}
                Check(exact==200,"User bank exact retrieval "+exact+"/200");Check(noisy==200,"User bank OCR spacing + options retrieval "+noisy+"/200");
                string[] fuzzyQueries={"电流速段保护是？","电流速断保护","电力系统出现振荡时，哪种保护可能误动？","两相两继电器的不完全星形接法能反应什么故障？"};int[] targets={1,1,2,0};
                for(int i=0;i<fuzzyQueries.Length;i++){var fuzzyHits=realMatcher.Search(fuzzyQueries[i],5);Check(fuzzyHits.Count>0&&fuzzyHits[0].Question==realQs[targets[i]]&&fuzzyHits[0].Score>=45,"User bank fuzzy query: "+fuzzyQueries[i]);}
                string image=Path.Combine(output,"ocr-fixture.png");
                using(var b=new Bitmap(2000,550)){using(var g=Graphics.FromImage(b)){g.Clear(Color.White);using(var f=new Font("Microsoft YaHei UI",24))g.DrawString("2. 电流速断保护是（ ）。\nA. 能保护线路全长；\nB. 不能保护线路全长；\nC. 有时能保护线路全长；",f,Brushes.Black,new RectangleF(35,35,1900,480));}b.Save(image,ImageFormat.Png);}
                using(var worker=new OcrWorker()){string text=worker.Recognize(image).GetAwaiter().GetResult();report.AppendLine("OCR TEXT: "+text.Replace("\n"," / "));var hits=realMatcher.Search(text,5);Check(Matcher.Normalize(text).Contains("电流") && hits.Count>0 && hits[0].Question==realQs[1] && hits[0].Question.Answer=="B","Real Windows OCR to user bank answer B");}
                using(var form=new MainForm()){
                    form.Show();Application.DoEvents();form.PerformLayout();var field=Controls(form).OfType<TextBox>().First(c=>c.Name=="ManualQuery");field.Text="电流速断保护是";
                    
                    Controls(form).OfType<Button>().First(b=>b.Text=="搜索").PerformClick();
                    typeof(MainForm).GetMethod("ManualSearch",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).Invoke(form,null);
                    report.AppendLine("UI DETAIL: "+Controls(form).OfType<RichTextBox>().Single().Text.Replace("\n"," / "));Check(Controls(form).OfType<RichTextBox>().Single().Text.Contains("不能保护线路全长"),"Desktop form manual search shows answer content");
                    ShowWindow(form.Handle,9);form.TopMost=true;form.BringToFront();form.Refresh();Application.DoEvents();Thread.Sleep(200);
                    using(var b=new Bitmap(form.Width,form.Height)){using(var g=Graphics.FromImage(b))g.CopyFromScreen(form.Location,Point.Empty,b.Size);b.Save(Path.Combine(output,"app-preview.png"),ImageFormat.Png);}
                    form.TopMost=false;var ready=Controls(form).OfType<Button>().First(b=>b.Text=="开始实时识别");var timeout=System.Diagnostics.Stopwatch.StartNew();while(!ready.Enabled&&timeout.ElapsedMilliseconds<22000){Application.DoEvents();Thread.Sleep(20);}
                    using(var fixture=new Form{Text="屏幕识别验证",StartPosition=FormStartPosition.Manual,Location=new Point(30,40),ClientSize=new Size(960,320),TopMost=true}) {
                        var picture=new PictureBox{Dock=DockStyle.Fill,SizeMode=PictureBoxSizeMode.Normal};fixture.Controls.Add(picture);fixture.Show();ShowWindow(fixture.Handle,9);form.WindowState=FormWindowState.Minimized;
                        var regionField=typeof(MainForm).GetField("region",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance);
                        var scan=typeof(MainForm).GetMethod("Scan",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance);
                        string[] texts={"2. 电流速断保护是（ ）。\nA. 能保护线路全长；\nB. 不能保护线路全长；\nC. 有时能保护线路全长；","3. 电力系统发生振荡时，（ ）可能会发生误动。\nA. 电流差动保护；\nB. 零序电流速断保护；\nC. 电流速断保护。",""};
                        for(int step=0;step<texts.Length;step++) {
                            using(var screen=new Bitmap(960,320)){using(var g=Graphics.FromImage(screen)){g.Clear(Color.White);using(var f=new Font("Microsoft YaHei UI",20))g.DrawString(texts[step],f,Brushes.Black,new RectangleF(20,20,920,280));}if(picture.Image!=null)picture.Image.Dispose();picture.Image=new Bitmap(screen);}
                            fixture.Refresh();Application.DoEvents();Thread.Sleep(200);using(var snap=new Bitmap(960,320)){using(var g=Graphics.FromImage(snap))g.CopyFromScreen(fixture.PointToScreen(Point.Empty),Point.Empty,snap.Size);snap.Save(Path.Combine(output,"live-"+step+".png"),ImageFormat.Png);}regionField.SetValue(form,fixture.RectangleToScreen(fixture.ClientRectangle));
                            var task=(Task)scan.Invoke(form,new object[]{true});timeout.Restart();while(!task.IsCompleted&&timeout.ElapsedMilliseconds<25000){Application.DoEvents();Thread.Sleep(20);}if(!task.IsCompleted)throw new Exception("屏幕识别测试超时");task.GetAwaiter().GetResult();
                            report.AppendLine("OCR TIMING: "+((OcrWorker)typeof(MainForm).GetField("ocr",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).GetValue(form)).LastTiming);string shown=Controls(form).OfType<RichTextBox>().Single().Text;report.AppendLine("LIVE STATUS: "+string.Join(" / ",Controls(form).OfType<Label>().Select(l=>l.Text)));report.AppendLine("LIVE OCR: "+string.Join(" / ",Controls(form).OfType<TextBox>().Where(t=>t.Multiline).Select(t=>t.Text)));report.AppendLine("LIVE STEP "+step+": "+shown.Replace("\n"," / "));
                            Check(step==0?shown.Contains("不能保护线路全长"):step==1?shown.Contains("发生振荡") && !shown.Contains("对应选项内容：不能保护线路全长"):shown=="","Live screen capture step "+step+" and stale-answer clearing");
                        }
                        if(picture.Image!=null)picture.Image.Dispose();
                    }
                    string cache=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"题库缓存.json");byte[] previousCache=File.Exists(cache)?File.ReadAllBytes(cache):null;
                    try {
                        typeof(MainForm).GetMethod("SaveBank",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).Invoke(form,null);
                        using(var restored=new MainForm()) {
                            typeof(MainForm).GetMethod("LoadDefault",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).Invoke(restored,null);
                            var restoredMatcher=(Matcher)typeof(MainForm).GetField("matcher",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).GetValue(restored);
                            Check(restoredMatcher.Count==200 && restoredMatcher.Search("电流速断保护是",5)[0].Question.Answer=="B","Persisted bank restores questions, options and answers");
                        }
                    } finally {if(previousCache==null)File.Delete(cache);else File.WriteAllBytes(cache,previousCache);}
                }
                using(var mapping=new MappingDialog(realSheets)){mapping.CreateControl();Check(mapping.Controls.Count>0,"Excel mapping dialog builds");}
            }catch(Exception ex){failed++;report.AppendLine("ERROR "+ex.ToString());}
            report.AppendLine("Passed: "+passed+"; Failed: "+failed);File.WriteAllText(Path.Combine(output,"test-report.txt"),report.ToString(),new UTF8Encoding(true));return failed==0?0:1;
        }
    }
}
