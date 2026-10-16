using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace ScreenBank {
    static class Program {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
        [STAThread] static void Main(string[] args) {
            try {SetProcessDpiAwarenessContext(new IntPtr(-4));}catch {SetProcessDPIAware();}
            int threads,io;System.Threading.ThreadPool.GetMinThreads(out threads,out io);System.Threading.ThreadPool.SetMinThreads(Math.Max(8,threads),Math.Max(8,io));
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            if(args.Length>0 && args[0]=="--self-test") { Environment.Exit(SelfTest.Run(args.Skip(1).ToArray()));return; }
            Application.Run(new MainForm());
        }
    }
    public sealed class OcrWorker : IDisposable {
        Process process; readonly JavaScriptSerializer json=new JavaScriptSerializer();
        public int MaxDimension=2600; public string Language,LastTiming;
        async Task<Dictionary<string,object>> Read() {
            var task=process.StandardOutput.ReadLineAsync();
            if(await Task.WhenAny(task,Task.Delay(20000))!=task) {Dispose();throw new Exception("OCR 响应超时，请重试。");}
            var line=await task; if(line==null)throw new Exception("OCR 进程退出。请检查 ocr-worker.ps1 是否与程序放在同一文件夹。");
            var response=json.Deserialize<Dictionary<string,object>>(line);
            if(response.ContainsKey("error"))throw new Exception(response["error"].ToString());return response;
        }
        public async Task Initialize() {
            if(process!=null && !process.HasExited)return;
            Dispose();string script=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"ocr-worker.ps1");
            if(!File.Exists(script))throw new Exception("缺少 ocr-worker.ps1，请完整解压软件压缩包。");
            var info=new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"WindowsPowerShell\\v1.0\\powershell.exe"),"-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File \""+script+"\"") {
                UseShellExecute=false,CreateNoWindow=true,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,StandardOutputEncoding=new UTF8Encoding(false),StandardErrorEncoding=new UTF8Encoding(false)
            };
            process=Process.Start(info);process.StandardInput.AutoFlush=true;
            process.ErrorDataReceived+=(s,e)=>{};process.BeginErrorReadLine();
            var ready=await Read();Language=ready["language"].ToString();MaxDimension=Convert.ToInt32(ready["maxDimension"]);
        }
        public async Task<string> Recognize(string path) {
            try {var watch=Stopwatch.StartNew();await Initialize();long init=watch.ElapsedMilliseconds;await process.StandardInput.WriteLineAsync(Regex.Replace(json.Serialize(new {path=path}), @"[^\x00-\x7F]", m=>"\\u"+((int)m.Value[0]).ToString("x4")));long sent=watch.ElapsedMilliseconds;var result=await Read();LastTiming="init="+init+" send="+(sent-init)+" read="+(watch.ElapsedMilliseconds-sent)+" worker="+(result.ContainsKey("elapsedMs")?result["elapsedMs"].ToString():"?");return result["text"].ToString();}
            catch {Dispose();throw;}
        }
        public void Dispose() { if(process!=null){try{if(!process.HasExited)process.Kill();}catch{}process.Dispose();process=null;} }
    }
    public sealed class Selector : Form {
        Point start,finish; bool dragging; Bitmap desktop;
        public Rectangle Selected;
        public Selector() {
            FormBorderStyle=FormBorderStyle.None;StartPosition=FormStartPosition.Manual;Bounds=SystemInformation.VirtualScreen;TopMost=true;ShowInTaskbar=false;DoubleBuffered=true;Cursor=Cursors.Cross;KeyPreview=true;
            desktop=new Bitmap(Width,Height);using(var g=Graphics.FromImage(desktop))g.CopyFromScreen(Bounds.Location,Point.Empty,Bounds.Size);
            KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Escape){DialogResult=DialogResult.Cancel;Close();}};
            MouseDown+=(s,e)=>{if(e.Button==MouseButtons.Right){DialogResult=DialogResult.Cancel;Close();}else if(e.Button==MouseButtons.Left){start=finish=e.Location;dragging=true;Invalidate();}};
            MouseMove+=(s,e)=>{if(dragging){finish=e.Location;Invalidate();}};
            MouseUp+=(s,e)=>{if(!dragging||e.Button!=MouseButtons.Left)return;finish=e.Location;var r=Selection();if(r.Width<30||r.Height<30){dragging=false;Invalidate();return;}Selected=new Rectangle(r.X+Left,r.Y+Top,r.Width,r.Height);DialogResult=DialogResult.OK;Close();};
        }
        Rectangle Selection(){return Rectangle.FromLTRB(Math.Min(start.X,finish.X),Math.Min(start.Y,finish.Y),Math.Max(start.X,finish.X),Math.Max(start.Y,finish.Y));}
        protected override void OnPaint(PaintEventArgs e){
            e.Graphics.DrawImageUnscaled(desktop,0,0);using(var shade=new SolidBrush(Color.FromArgb(145,0,0,0)))e.Graphics.FillRectangle(shade,ClientRectangle);
            if(dragging){var r=Selection();e.Graphics.DrawImage(desktop,r,r,GraphicsUnit.Pixel);using(var pen=new Pen(Color.Turquoise,3))e.Graphics.DrawRectangle(pen,r);}
            string help="拖动框选投屏中的题目和选项  ·  Esc / 右键取消";
            Point p=PointToClient(Cursor.Position);p.Offset(18,22);p.X=Math.Max(8,Math.Min(p.X,Width-550));p.Y=Math.Max(8,Math.Min(p.Y,Height-70));
            using(var b=new SolidBrush(Color.FromArgb(225,18,34,44)))e.Graphics.FillRectangle(b,p.X,p.Y,540,46);
            TextRenderer.DrawText(e.Graphics,help,new Font("Microsoft YaHei UI",12),new Rectangle(p.X+12,p.Y+10,516,28),Color.White);
        }
        protected override void Dispose(bool disposing){if(disposing&&desktop!=null){desktop.Dispose();desktop=null;}base.Dispose(disposing);}
    }
    public sealed class MappingDialog : Form {
        ComboBox sheet=new ComboBox(), row=new ComboBox(); ComboBox[] fields=new ComboBox[10];DataGridView preview=new DataGridView();
        List<Sheet> sheets;public Sheet Chosen;public int HeaderIndex;public int[] Map;
        public MappingDialog(List<Sheet> data) {
            sheets=data;Text="确认 Excel 列对应关系";Font=new Font("Microsoft YaHei UI",10);Size=new Size(950,620);MinimumSize=new Size(930,560);StartPosition=FormStartPosition.CenterParent;
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=4,Padding=new Padding(16)};Controls.Add(root);
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,44));root.RowStyles.Add(new RowStyle(SizeType.Absolute,130));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            var top=new FlowLayoutPanel{Dock=DockStyle.Fill};root.Controls.Add(top,0,0);
            top.Controls.Add(new Label{Text="工作表",AutoSize=true,Margin=new Padding(0,8,8,0)});sheet.Width=210;sheet.DropDownStyle=ComboBoxStyle.DropDownList;top.Controls.Add(sheet);
            top.Controls.Add(new Label{Text="表头行",AutoSize=true,Margin=new Padding(20,8,8,0)});row.Width=120;row.DropDownStyle=ComboBoxStyle.DropDownList;top.Controls.Add(row);
            var mappings=new FlowLayoutPanel{Dock=DockStyle.Fill};root.Controls.Add(mappings,0,1);
            string[] labels={"题目（必选）","答案","解析","选项 A","选项 B","选项 C","选项 D","选项 E","选项 F","选项 G"};
            for(int i=0;i<fields.Length;i++){var box=new FlowLayoutPanel{FlowDirection=FlowDirection.TopDown,Width=165,Height=60};box.Controls.Add(new Label{Text=labels[i],AutoSize=true});fields[i]=new ComboBox{Width=156,DropDownStyle=ComboBoxStyle.DropDownList};box.Controls.Add(fields[i]);mappings.Controls.Add(box);}
            preview.Dock=DockStyle.Fill;preview.ReadOnly=true;preview.AllowUserToAddRows=false;preview.RowHeadersVisible=false;preview.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.AllCells;root.Controls.Add(preview,0,2);
            var bottom=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft};root.Controls.Add(bottom,0,3);
            var ok=new Button{Text="导入当前工作表",Width=170,Height=34};ok.Click+=(s,e)=>{Map=fields.Select(f=>f.SelectedIndex-1).ToArray();if(Map[0]<0){MessageBox.Show(this,"请选择题目列。");return;}Chosen=sheets[sheet.SelectedIndex];HeaderIndex=row.SelectedIndex;DialogResult=DialogResult.OK;};bottom.Controls.Add(ok);
            bottom.Controls.Add(new Label{Text="可重复导入其他工作表或文件，题目将合并。",AutoSize=true,Margin=new Padding(8,8,12,0)});
            sheet.SelectedIndexChanged+=(s,e)=>{row.Items.Clear();var current=sheets[sheet.SelectedIndex];for(int i=0;i<Math.Min(20,current.Rows.Count);i++)row.Items.Add("第 "+(i+1)+" 行");row.SelectedIndex=BankFile.HeaderRow(current);};
            row.SelectedIndexChanged+=(s,e)=>RefreshFields();foreach(var item in data)sheet.Items.Add(item.Name);sheet.SelectedIndex=0;
        }
        void RefreshFields(){
            if(sheet.SelectedIndex<0||row.SelectedIndex<0)return;var current=sheets[sheet.SelectedIndex];var header=current.Rows[row.SelectedIndex];var guess=BankFile.Guess(header);
            for(int i=0;i<fields.Length;i++){fields[i].Items.Clear();fields[i].Items.Add("（不导入）");for(int j=0;j<header.Length;j++)fields[i].Items.Add((j+1)+" · "+header[j]);fields[i].SelectedIndex=guess[i]+1;}
            preview.Columns.Clear();for(int j=0;j<header.Length;j++)preview.Columns.Add("c"+j,string.IsNullOrWhiteSpace(header[j])?"列 "+(j+1):header[j]);
            for(int i=row.SelectedIndex+1;i<Math.Min(current.Rows.Count,row.SelectedIndex+6);i++){var values=Enumerable.Repeat("",header.Length).ToArray();Array.Copy(current.Rows[i],values,Math.Min(values.Length,current.Rows[i].Length));preview.Rows.Add(values);}
        }
    }
    public sealed class MainForm : Form {
        readonly string basePath=AppDomain.CurrentDomain.BaseDirectory;
        readonly OcrWorker ocr=new OcrWorker(); readonly Timer timer=new Timer(); readonly List<Question> bank=new List<Question>();
        Matcher matcher;Rectangle region;bool running,busy,closing,updatingResults;int generation;string lastHash="",lastText="";string capturePath;StringBuilder rich;
        Label status=new Label(),bankLabel=new Label(),regionLabel=new Label(); Button run=new Button(),once=new Button(),choose=new Button(),import=new Button();
        TextBox recognized=new TextBox(),manual=new TextBox();RichTextBox detail=new RichTextBox();DataGridView results=new DataGridView();PictureBox preview=new PictureBox();
        NumericUpDown interval=new NumericUpDown(),threshold=new NumericUpDown();CheckBox topmost=new CheckBox(),enlarge=new CheckBox();
        public MainForm(){
            Text="屏幕题库助手 · 本地实时匹配";Font=new Font("Microsoft YaHei UI",10);BackColor=Color.FromArgb(244,247,249);Size=new Size(1130,810);MinimumSize=new Size(1000,720);StartPosition=FormStartPosition.CenterScreen;
            capturePath=Path.Combine(basePath,".cache-"+Guid.NewGuid().ToString("N")+".png");BuildUi();run.Enabled=false;once.Enabled=false;
            timer.Interval=1200;timer.Tick+=async(s,e)=>await Scan(false);
            Shown+=async(s,e)=>{LoadDefault();try{Status("正在初始化中文 OCR…");await ocr.Initialize();if(!closing)Status("就绪 · OCR "+ocr.Language+" · 框选区域后开始识别");}catch(Exception ex){if(!closing)Status(ex.Message);}finally{if(!closing){run.Enabled=true;once.Enabled=true;}}};
            FormClosing+=(s,e)=>{closing=true;running=false;generation++;timer.Stop();ocr.Dispose();if(preview.Image!=null)preview.Image.Dispose();try{File.Delete(capturePath);}catch{}};
        }
        Button Button(string text,int width){return new Button{Text=text,Width=width,Height=36,FlatStyle=FlatStyle.Flat,BackColor=Color.White,Margin=new Padding(0,0,9,0)};}
        void BuildUi(){
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=5,Padding=new Padding(20)};Controls.Add(root);
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,65));root.RowStyles.Add(new RowStyle(SizeType.Absolute,46));root.RowStyles.Add(new RowStyle(SizeType.Absolute,45));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,43));
            var title=new Panel{Dock=DockStyle.Fill};root.Controls.Add(title,0,0);title.Controls.Add(new Label{Text="屏幕题库助手",Font=new Font(Font.FontFamily,21,FontStyle.Bold),ForeColor=Color.FromArgb(18,71,82),AutoSize=true});bankLabel.Text="本地 OCR  /  Excel 题库  /  相似题检索";bankLabel.AutoSize=true;bankLabel.Location=new Point(2,42);title.Controls.Add(bankLabel);
            var tools=new FlowLayoutPanel{Dock=DockStyle.Fill};root.Controls.Add(tools,0,1);
            import=Button("导入 Excel",126);import.Click+=async(s,e)=>await Import();tools.Controls.Add(import);
            choose=Button("框选屏幕区域",145);choose.Click+=async(s,e)=>await Choose();tools.Controls.Add(choose);
            run=Button("开始实时识别",150);run.BackColor=Color.FromArgb(18,108,117);run.ForeColor=Color.White;run.Click+=async(s,e)=>{if(running){Stop();Status("已暂停 · 可手动搜索或重新框选");}else{if(!Ready())return;running=true;run.Text="暂停识别";lastHash="";generation++;timer.Start();await Scan(true);}};tools.Controls.Add(run);
            once=Button("识别一次",118);once.Click+=async(s,e)=>{if(Ready())await Scan(true);};tools.Controls.Add(once);
            var clear=Button("清空题库",108);clear.Click+=(s,e)=>{Stop();bank.Clear();matcher=null;ClearResults();BankStatus();SaveBank();Status("题库已清空，请导入 Excel。");};tools.Controls.Add(clear);
            topmost.Text="窗口置顶";topmost.AutoSize=true;topmost.Margin=new Padding(10,7,0,0);topmost.CheckedChanged+=(s,e)=>TopMost=topmost.Checked;tools.Controls.Add(topmost);
            var prefs=new FlowLayoutPanel{Dock=DockStyle.Fill};root.Controls.Add(prefs,0,2);
            prefs.Controls.Add(new Label{Text="间隔（秒）",AutoSize=true,Margin=new Padding(0,8,6,0)});interval.Minimum=0.5M;interval.Maximum=10;interval.DecimalPlaces=1;interval.Increment=0.5M;interval.Value=1.2M;interval.Width=70;interval.ValueChanged+=(s,e)=>timer.Interval=(int)(interval.Value*1000);prefs.Controls.Add(interval);
            prefs.Controls.Add(new Label{Text="匹配分数下限",AutoSize=true,Margin=new Padding(18,8,6,0)});threshold.Minimum=20;threshold.Maximum=95;threshold.Value=45;threshold.Width=65;threshold.ValueChanged+=(s,e)=>{if(matcher!=null)DisplayHits(matcher.Search(lastText,5));};prefs.Controls.Add(threshold);
            enlarge.Text="小字放大";enlarge.AutoSize=true;enlarge.Margin=new Padding(15,6,0,0);enlarge.CheckedChanged+=(s,e)=>lastHash="";prefs.Controls.Add(enlarge);
            regionLabel.AutoSize=true;regionLabel.Text="区域：未选择";regionLabel.Margin=new Padding(18,8,0,0);prefs.Controls.Add(regionLabel);
            var split=new SplitContainer{Size=new Size(1000,500),Dock=DockStyle.Fill,SplitterDistance=520,Panel1MinSize=390,Panel2MinSize=380};root.Controls.Add(split,0,3);
            var left=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=6,Padding=new Padding(0,0,12,0)};split.Panel1.Controls.Add(left);
            left.RowStyles.Add(new RowStyle(SizeType.Absolute,30));left.RowStyles.Add(new RowStyle(SizeType.Percent,46));left.RowStyles.Add(new RowStyle(SizeType.Absolute,28));left.RowStyles.Add(new RowStyle(SizeType.Percent,31));left.RowStyles.Add(new RowStyle(SizeType.Absolute,72));left.RowStyles.Add(new RowStyle(SizeType.Percent,23));
            left.Controls.Add(new Label{Text="相似题候选  ·  点击查看题库答案",AutoSize=true},0,0);
            results.Dock=DockStyle.Fill;results.BackgroundColor=Color.White;results.BorderStyle=BorderStyle.None;results.ReadOnly=true;results.AllowUserToAddRows=false;results.AllowUserToDeleteRows=false;results.AllowUserToResizeRows=false;results.RowHeadersVisible=false;results.SelectionMode=DataGridViewSelectionMode.FullRowSelect;results.MultiSelect=false;results.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;results.AutoGenerateColumns=false;
            results.Columns.Add(new DataGridViewTextBoxColumn{HeaderText="分数",Name="score",FillWeight=20});results.Columns.Add(new DataGridViewTextBoxColumn{HeaderText="题目",Name="question",FillWeight=80});results.RowTemplate.Height=37;results.DefaultCellStyle.SelectionBackColor=Color.FromArgb(219,240,241);results.DefaultCellStyle.SelectionForeColor=Color.Black;results.SelectionChanged+=(s,e)=>{if(!updatingResults&&results.SelectedRows.Count>0)ShowQuestion((Hit)results.SelectedRows[0].Tag);};left.Controls.Add(results,0,1);
            left.Controls.Add(new Label{Text="本次 OCR 文字（实时模式下自动更新）",AutoSize=true},0,2);recognized.Multiline=true;recognized.ReadOnly=true;recognized.ScrollBars=ScrollBars.Vertical;recognized.Dock=DockStyle.Fill;recognized.BackColor=Color.White;left.Controls.Add(recognized,0,3);
            var search=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=2,Padding=new Padding(0,8,0,5)};search.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));search.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,95));search.Controls.Add(new Label{Text="手动输入关键词（搜索时暂停实时识别）",AutoSize=true},0,0);search.SetColumnSpan(search.Controls[0],2);manual.Name="ManualQuery";manual.Dock=DockStyle.Fill;search.Controls.Add(manual,0,1);var find=Button("搜索",88);find.Height=29;find.Click+=(s,e)=>ManualSearch();search.Controls.Add(find,1,1);manual.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Enter){e.SuppressKeyPress=true;ManualSearch();}};left.Controls.Add(search,0,4);
            preview.Dock=DockStyle.Fill;preview.SizeMode=PictureBoxSizeMode.Zoom;preview.BackColor=Color.FromArgb(228,234,238);left.Controls.Add(preview,0,5);
            var right=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(14,0,0,0)};right.RowStyles.Add(new RowStyle(SizeType.Absolute,30));right.RowStyles.Add(new RowStyle(SizeType.Percent,100));right.RowStyles.Add(new RowStyle(SizeType.Absolute,52));split.Panel2.Controls.Add(right);
            right.Controls.Add(new Label{Text="对应题目与题库答案",AutoSize=true,ForeColor=Color.FromArgb(18,71,82)},0,0);detail.Dock=DockStyle.Fill;detail.ReadOnly=true;detail.BorderStyle=BorderStyle.None;detail.BackColor=Color.White;detail.Font=new Font(Font.FontFamily,12);detail.Text="导入题库 → 框选手机投屏区域 → 开始实时识别\n\n建议框选完整题干及选项，提高相似题区分效果。";right.Controls.Add(detail,0,1);
            var copy=Button("复制当前题目与答案",210);copy.Click+=(s,e)=>{if(results.SelectedRows.Count>0)try{Clipboard.SetText(detail.Text);Status("已复制当前题目与答案。");}catch(Exception ex){Status(ex.Message);}};copy.Margin=new Padding(0,12,0,0);right.Controls.Add(copy,0,2);
            status.Dock=DockStyle.Fill;status.ForeColor=Color.FromArgb(56,76,86);status.Padding=new Padding(0,12,0,0);status.AutoEllipsis=true;root.Controls.Add(status,0,4);
        }
        void Status(string text){status.Text=text;}
        void BankStatus(){bankLabel.Text="已导入 "+bank.Count+" 道题  ·  本地中文 OCR  ·  Excel / CSV";}
        void Stop(){running=false;timer.Stop();run.Text="开始实时识别";generation++;lastHash="";}
        void ClearResults(){results.Rows.Clear();detail.Clear();}
        bool Ready(){if(matcher==null||matcher.Count==0){Status("请先导入题库。");return false;}if(region.IsEmpty){Status("请先框选手机投屏中的题目区域。");return false;}return true;}
        void LoadDefault(){
            try{string saved=Path.Combine(basePath,"题库缓存.json");if(File.Exists(saved)){var serializer=new JavaScriptSerializer{MaxJsonLength=100*1024*1024};bank.AddRange(serializer.Deserialize<List<Question>>(File.ReadAllText(saved,Encoding.UTF8)));}else{string file=Path.Combine(basePath,"电气专业试题库（第六周）.xlsx");if(!File.Exists(file))return;foreach(var sh in BankFile.Read(file)){int h=BankFile.HeaderRow(sh);var map=BankFile.Guess(sh.Rows[h]);if(map[0]>=0)bank.AddRange(BankFile.Convert(sh,h,map,file));}}matcher=new Matcher(bank);BankStatus();}catch(Exception ex){Status("默认题库加载失败："+ex.Message);}
        }
        void SaveBank(){
            try{var serializer=new JavaScriptSerializer{MaxJsonLength=100*1024*1024};string path=Path.Combine(basePath,"题库缓存.json"),temporary=path+".new";var saved=bank.Select(q=>new {q.Text,q.Answer,q.Explanation,q.Source,q.Options}).ToList();File.WriteAllText(temporary,serializer.Serialize(saved),new UTF8Encoding(true));if(File.Exists(path))File.Replace(temporary,path,null);else File.Move(temporary,path);}
            catch(Exception ex){MessageBox.Show(this,"题库在本次运行中已更新，但未能保存缓存："+ex.Message,"题库保存",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        }
        async Task Import(){
            using(var dialog=new OpenFileDialog{Filter="Excel 题库 (*.xlsx)|*.xlsx|CSV / TSV (*.csv;*.tsv)|*.csv;*.tsv",Title="选择题库（可重复导入合并）"}) {
                if(dialog.ShowDialog(this)!=DialogResult.OK)return;Stop();import.Enabled=false;
                try{Status("正在读取题库…");var sheets=await Task.Run(()=>BankFile.Read(dialog.FileName));if(closing)return;using(var mapping=new MappingDialog(sheets))if(mapping.ShowDialog(this)==DialogResult.OK){var added=BankFile.Convert(mapping.Chosen,mapping.HeaderIndex,mapping.Map,dialog.FileName);bank.AddRange(added);matcher=await Task.Run(()=>new Matcher(bank));if(closing)return;SaveBank();ClearResults();lastText="";BankStatus();Status("新增 "+added.Count+" 道题。总计 "+bank.Count+" 道题。");}else Status("已取消导入。");}
                catch(Exception ex){if(!closing)Status("导入失败："+ex.Message);}finally{if(!closing)import.Enabled=true;}
            }
        }
        async Task Choose(){
            if(busy){Status("请等待当前识别完成，再重新框选。");return;}Stop();WindowState=FormWindowState.Minimized;await Task.Delay(280);if(closing)return;
            try{using(var selector=new Selector())if(selector.ShowDialog()==DialogResult.OK){region=selector.Selected;regionLabel.Text=string.Format("区域：{0}, {1}  ·  {2} × {3}",region.X,region.Y,region.Width,region.Height);recognized.Clear();ClearResults();if(preview.Image!=null){preview.Image.Dispose();preview.Image=null;}Status("区域已选择，把本窗口放在区域外后开始识别。");}}
            catch(Exception ex){Status("框选失败："+ex.Message);}finally{if(!closing){WindowState=FormWindowState.Normal;Activate();}}
        }
        void ManualSearch(){Stop();if(matcher==null){Status("请先导入题库。");return;}lastText=manual.Text;DisplayHits(matcher.Search(lastText,5));Status("手动搜索 · 实时识别已暂停");}
        async Task Scan(bool force){
            if(busy||closing||matcher==null||region.IsEmpty)return;
            if(!SystemInformation.VirtualScreen.Contains(region)){Stop();ClearResults();Status("显示器布局已变化，请重新框选。");return;}
            if(WindowState!=FormWindowState.Minimized && Visible && Bounds.IntersectsWith(region)){lastHash="";ClearResults();recognized.Clear();Status("软件窗口遮挡了识别区域，请将窗口移开或最小化。");return;}
            busy=true;once.Enabled=false;choose.Enabled=false;int version=generation;
            try{
                using(var screenshot=new Bitmap(region.Width,region.Height,PixelFormat.Format32bppArgb)) {
                    using(var g=Graphics.FromImage(screenshot))g.CopyFromScreen(region.Location,Point.Empty,region.Size);
                    int max=Math.Min(ocr.MaxDimension,2000);double scale=Math.Min(enlarge.Checked?2.0:1.0,Math.Min(max/(double)screenshot.Width,max/(double)screenshot.Height));
                    using(var scaled=new Bitmap(Math.Max(1,(int)(screenshot.Width*scale)),Math.Max(1,(int)(screenshot.Height*scale)))) {
                        using(var g=Graphics.FromImage(scaled)){g.InterpolationMode=System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;g.DrawImage(screenshot,new Rectangle(Point.Empty,scaled.Size));}
                        byte[] bytes;using(var memory=new MemoryStream()){scaled.Save(memory,ImageFormat.Png);bytes=memory.ToArray();}
                        string hash;using(var sha=SHA256.Create())hash=Convert.ToBase64String(sha.ComputeHash(bytes));
                        if(!force&&hash==lastHash)return;
                        lastHash=hash;ClearResults();recognized.Clear();lastText="";
                        var old=preview.Image;preview.Image=new Bitmap(screenshot);if(old!=null)old.Dispose();
                        File.WriteAllBytes(capturePath,bytes);
                    }
                }
                Status("正在识别…");var watch=Stopwatch.StartNew();string text=await ocr.Recognize(capturePath);
                if(closing||version!=generation)return;recognized.Text=text;lastText=text;
                if(string.IsNullOrWhiteSpace(text)){Status("未识别到文字，请调整区域、投屏大小或清晰度。");return;}
                long readAt=watch.ElapsedMilliseconds;var activeMatcher=matcher;var hits=await Task.Run(()=>activeMatcher.Search(text,5));long matchAt=watch.ElapsedMilliseconds;if(closing||version!=generation)return;DisplayHits(hits);ocr.LastTiming+=" match="+(matchAt-readAt)+" display="+(watch.ElapsedMilliseconds-matchAt);
                int shown=results.Rows.Count;Status(DateTime.Now.ToString("HH:mm:ss")+" · "+watch.ElapsedMilliseconds+" ms · "+(shown==0?"无可靠候选，尝试只框选题干或降低下限":shown+" 个候选 · 分数表示文字相似程度，请核对题干和选项"));
            }catch(Exception ex){lastHash="";if(!closing){ClearResults();recognized.Clear();Status("识别失败："+ex.Message);Stop();}}
            finally{busy=false;if(!closing){once.Enabled=true;choose.Enabled=true;}try{File.Delete(capturePath);}catch{}}
        }
        void DisplayHits(List<Hit> hits){
            var visible=hits.Where(h=>h.Score>=(double)threshold.Value).ToList();updatingResults=true;
            try{ClearResults();foreach(var h in visible){int i=results.Rows.Add(h.Score.ToString("F0"),h.Question.Text);results.Rows[i].Tag=h;}if(visible.Count>0){results.ClearSelection();results.Rows[0].Selected=true;}}finally{updatingResults=false;}
            if(visible.Count>0)ShowQuestion(visible[0]);else detail.Text="未找到达到分数下限的题目。\n\n可以检查 OCR 文字、缩小框选范围，或手动输入一段题干。";
        }
        void Append(string text,Color color,bool bold){
            int index=color.R==18?1:color.R==30?2:color.R==15?3:color.R==110?4:5;
            rich.Append("{\\cf").Append(index).Append(bold?"\\b ":"\\b0 ");
            foreach(char c in text){if(c=='\\'||c=='{'||c=='}')rich.Append('\\').Append(c);else if(c=='\n')rich.Append("\\par ");else if(c=='\r')continue;else if(c=='\t')rich.Append("\\tab ");else if(c>=128)rich.Append("\\u").Append((short)c).Append('?');else if(c>=32)rich.Append(c);}
            rich.Append('}');
        }
        void ShowQuestion(Hit hit){
            if(hit==null)return;var q=hit.Question;rich=new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Microsoft YaHei UI;}}{\colortbl;\red18\green108\blue117;\red30\green45\blue52;\red15\green115\blue85;\red110\green121\blue130;\red47\green62\blue68;}\viewkind4\uc1\pard\f0\fs24 ");Append("匹配分数 "+hit.Score.ToString("F0")+" / 100\n\n",Color.FromArgb(18,108,117),true);Append(q.Text+"\n\n",Color.FromArgb(30,45,52),true);
            for(int i=0;i<q.Options.Length;i++)if(!string.IsNullOrWhiteSpace(q.Options[i]))Append(((char)('A'+i))+". "+q.Options[i]+"\n",Color.FromArgb(47,62,68),false);
            Append("\n题库答案："+(string.IsNullOrWhiteSpace(q.Answer)?"（题库未提供）":q.Answer)+"\n",Color.FromArgb(15,115,85),true);
            string letters=Regex.Replace((q.Answer??"").ToUpperInvariant(),@"[\s,，、;；/|]","");
            if(Regex.IsMatch(letters,@"^[A-G]+$")){foreach(char letter in letters.Distinct()){int index=letter-'A';if(index<q.Options.Length&&!string.IsNullOrWhiteSpace(q.Options[index]))Append("对应选项内容："+q.Options[index]+"\n",Color.FromArgb(15,115,85),true);}}
            if(!string.IsNullOrWhiteSpace(q.Explanation))Append("\n解析\n"+q.Explanation+"\n",Color.FromArgb(47,62,68),false);
            Append("\n来源："+q.Source+"\n\n请按选项内容核对，投屏端的选项顺序可能不同。",Color.FromArgb(110,121,130),false);rich.Append('}');detail.Rtf=rich.ToString();detail.SelectionStart=0;detail.ScrollToCaret();
        }
    }
}
