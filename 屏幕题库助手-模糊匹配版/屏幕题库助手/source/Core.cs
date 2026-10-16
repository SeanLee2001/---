using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace ScreenBank {
    public class Question {
        public string Text, Answer, Explanation, Source;
        public string[] Options = new string[7];
        public string Normal;
        public string Fuzzy;
        public HashSet<string> Tokens;
        public override string ToString() { return Text; }
    }
    public class Hit { public Question Question; public double Score; }
    public class Sheet { public string Name; public List<string[]> Rows = new List<string[]>(); }
    public static class BankFile {
        public static readonly string[][] Aliases = {
            new [] { "题目", "题干", "试题", "问题", "题目内容", "试题内容", "question", "title", "stem" },
            new [] { "答案", "正确答案", "参考答案", "标准答案", "answer", "correctanswer" },
            new [] { "解析", "答案解析", "题目解析", "解答", "说明", "explanation", "analysis" },
            new [] { "a", "选项a", "a选项", "optiona" }, new [] { "b", "选项b", "b选项", "optionb" },
            new [] { "c", "选项c", "c选项", "optionc" }, new [] { "d", "选项d", "d选项", "optiond" },
            new [] { "e", "选项e", "e选项", "optione" }, new [] { "f", "选项f", "f选项", "optionf" }, new [] { "g", "选项g", "g选项", "optiong" }
        };
        public static string Header(string s) { return Regex.Replace((s ?? "").ToLowerInvariant(), @"[\s\p{P}\p{S}]", ""); }
        public static int[] Guess(string[] headers) {
            return Aliases.Select(a => Array.FindIndex(headers, h => a.Contains(Header(h)))).ToArray();
        }
        public static int HeaderRow(Sheet s) {
            int best = 0, score = -1;
            for (int i = 0; i < Math.Min(20, s.Rows.Count); i++) {
                var m = Guess(s.Rows[i]); int n = m.Count(x => x >= 0) + (m[0] >= 0 ? 5 : 0);
                if (n > score) { score = n; best = i; }
            }
            return best;
        }
        static string Get(string[] row, int col) { return col >= 0 && col < row.Length ? row[col].Trim() : ""; }
        public static List<Question> Convert(Sheet sheet, int header, int[] map, string file) {
            if (map[0] < 0) throw new Exception("请选择题目列。");
            var list = new List<Question>();
            for (int i = header + 1; i < sheet.Rows.Count; i++) {
                var r = sheet.Rows[i]; string q = Get(r, map[0]);
                if (string.IsNullOrWhiteSpace(q)) continue;
                var item = new Question { Text=q, Answer=Get(r,map[1]), Explanation=Get(r,map[2]), Source=Path.GetFileName(file)+" / "+sheet.Name+" / 第 "+(i+1)+" 行" };
                for (int j = 0; j < 7; j++) item.Options[j] = Get(r,map[j+3]);
                list.Add(item);
            }
            return list;
        }
        static XDocument Xml(ZipArchive z, string path) {
            var e = z.GetEntry(path);
            if (e == null) throw new Exception("Excel 缺少 "+path);
            if (e.Length > 150L*1024*1024) throw new Exception("工作表过大，请拆分题库。");
            using (var s = e.Open()) using (var reader = System.Xml.XmlReader.Create(s, new System.Xml.XmlReaderSettings { DtdProcessing=System.Xml.DtdProcessing.Prohibit, MaxCharactersInDocument=150L*1024*1024 })) return XDocument.Load(reader);
        }
        static int Column(string cell) {
            int n=0; foreach(char c in cell) { if (!char.IsLetter(c)) break; n=n*26+char.ToUpperInvariant(c)-'A'+1; }
            return n-1;
        }
        static string ResolvePart(string target) {
            var uri = new Uri(new Uri("http://package/xl/workbook.xml"), target.Replace('\\','/'));
            return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }
        public static List<Sheet> Read(string path) {
            if (new FileInfo(path).Length > 100L*1024*1024) throw new Exception("文件超过 100 MB，请拆分后导入。");
            var ext=Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xlsx") return Xlsx(path);
            if (ext == ".csv" || ext == ".tsv") return new List<Sheet> { new Sheet { Name="题库", Rows=Csv(File.ReadAllText(path,Encoding.UTF8), ext==".tsv"?'\t':',') } };
            throw new Exception("支持 .xlsx、UTF-8 CSV 和 TSV。旧版 .xls 请先在 Excel 中另存为 .xlsx。");
        }
        static List<Sheet> Xlsx(string path) {
            using(var z=ZipFile.OpenRead(path)) {
                XNamespace ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main", rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var shared = new List<string>();
                if (z.GetEntry("xl/sharedStrings.xml")!=null) shared=Xml(z,"xl/sharedStrings.xml").Descendants(ns+"si").Select(si=>string.Concat(si.Descendants(ns+"t").Select(t=>t.Value))).ToList();
                var relations=Xml(z,"xl/_rels/workbook.xml.rels").Root.Elements().ToDictionary(e=>(string)e.Attribute("Id"),e=>(string)e.Attribute("Target"));
                var result=new List<Sheet>();
                foreach(var sh in Xml(z,"xl/workbook.xml").Descendants(ns+"sheet")) {
                    string id=(string)sh.Attribute(rel+"id"); if(id==null || !relations.ContainsKey(id)) continue;
                    var sheet=new Sheet { Name=(string)sh.Attribute("name") };
                    foreach(var row in Xml(z,ResolvePart(relations[id])).Descendants(ns+"sheetData").Elements(ns+"row")) {
                        var cells=new SortedDictionary<int,string>(); int fallback=0;
                        foreach(var c in row.Elements(ns+"c")) {
                            string reference=(string)c.Attribute("r"); int col=reference==null?fallback:Column(reference); fallback=col+1;
                            if(col<0 || col>1023) throw new Exception("工作表超过 1024 列，请仅保留题库列。");
                            string type=(string)c.Attribute("t"), value=(string)c.Element(ns+"v")??"";
                            if(type=="s") { int index; value=int.TryParse(value,out index)&&index>=0&&index<shared.Count?shared[index]:""; }
                            else if(type=="inlineStr") value=string.Concat(c.Descendants(ns+"t").Select(t=>t.Value));
                            else if(c.Element(ns+"f")!=null && c.Element(ns+"v")==null) throw new Exception("发现没有缓存结果的公式，请在 Excel 中计算并保存，或粘贴为值后导入。");
                            cells[col]=value;
                        }
                        if(cells.Count==0) { sheet.Rows.Add(new string[0]); continue; }
                        var values=Enumerable.Repeat("",cells.Keys.Max()+1).ToArray(); foreach(var c in cells) values[c.Key]=c.Value;
                        sheet.Rows.Add(values);
                        if(sheet.Rows.Count>200000) throw new Exception("每个工作表最多 20 万行，请拆分题库。");
                    }
                    if(sheet.Rows.Count>0) result.Add(sheet);
                }
                if(result.Count==0) throw new Exception("Excel 没有可读取的工作表。");
                return result;
            }
        }
        public static List<string[]> Csv(string input,char separator) {
            var rows=new List<string[]>(); var row=new List<string>(); var b=new StringBuilder(); bool quoted=false;
            input=input.TrimStart('\uFEFF');
            for(int i=0;i<input.Length;i++) {
                char c=input[i];
                if(c=='"') { if(quoted && i+1<input.Length && input[i+1]=='"') { b.Append('"');i++; } else quoted=!quoted; }
                else if(!quoted && c==separator) { row.Add(b.ToString()); b.Clear(); }
                else if(!quoted && (c=='\r'||c=='\n')) { if(c=='\r'&&i+1<input.Length&&input[i+1]=='\n')i++; row.Add(b.ToString()); b.Clear();rows.Add(row.ToArray());row.Clear(); }
                else b.Append(c);
            }
            if(quoted) throw new Exception("CSV 引号没有闭合。");
            if(b.Length>0||row.Count>0) {row.Add(b.ToString());rows.Add(row.ToArray());}
            return rows;
        }
        public static void Template(string path) {
            XNamespace n="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var rows=new [] { new [] {"题目","A","B","C","D","答案","解析"},
                new [] {"中国的首都是哪里？","北京","上海","广州","深圳","A","中国的首都是北京。"},
                new [] {"下列哪个数字是质数？","4","6","7","9","C","7 只有 1 和自身两个正因数。"},
                new [] {"下列哪个数字不是质数？","2","3","4","5","C","4 可以被 2 整除。"},
                new [] {"水的化学式是什么？","H2O","CO2","O2","NaCl","A","水的化学式是 H2O。"} };
            using(var z=ZipFile.Open(path,ZipArchiveMode.Create)) {
                WriteEntry(z,"[Content_Types].xml","<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
                WriteEntry(z,"_rels/.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                WriteEntry(z,"xl/workbook.xml","<workbook xmlns=\""+n+"\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"题库\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                WriteEntry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                var data=new XElement(n+"sheetData");
                for(int i=0;i<rows.Length;i++) data.Add(new XElement(n+"row",new XAttribute("r",i+1),rows[i].Select((v,j)=>new XElement(n+"c",new XAttribute("r",((char)('A'+j)).ToString()+(i+1)),new XAttribute("t","inlineStr"),new XElement(n+"is",new XElement(n+"t",v))))));
                WriteEntry(z,"xl/worksheets/sheet1.xml",new XDocument(new XElement(n+"worksheet",data)).ToString());
            }
        }
        static void WriteEntry(ZipArchive z,string name,string value) { using(var w=new StreamWriter(z.CreateEntry(name).Open(),new UTF8Encoding(false)))w.Write(value); }
    }
    public class Matcher {
        readonly List<Question> questions;
        readonly Dictionary<string,List<int>> index=new Dictionary<string,List<int>>();
        readonly Dictionary<char,List<int>> charIndex=new Dictionary<char,List<int>>();
        public int Count { get {return questions.Count;} }
        public Matcher(List<Question> qs) {
            questions=qs;
            for(int i=0;i<qs.Count;i++) {
                var q=qs[i];q.Normal=Normalize(q.Text);q.Fuzzy=FuzzyNormalize(q.Text);q.Tokens=Tokens(q.Fuzzy);
                foreach(var t in q.Tokens) { List<int> ids; if(!index.TryGetValue(t,out ids))index[t]=ids=new List<int>();ids.Add(i); }
                foreach(char c in q.Fuzzy.Distinct()){List<int> ids;if(!charIndex.TryGetValue(c,out ids))charIndex[c]=ids=new List<int>();ids.Add(i);}
            }
        }
        public static string Normalize(string s) {
            s=(s??"").Normalize(NormalizationForm.FormKC).ToLowerInvariant();
            s=Regex.Replace(s,@"^\s*(?:第\s*)?\d+\s*(?:题\s*[:：.、]?|[.、)）:：])\s*", "");
            return Regex.Replace(s,@"[^\p{L}\p{Nd}]","");
        }
        static HashSet<string> Tokens(string s) { var set=new HashSet<string>(); if(s.Length==1)set.Add(s);for(int i=0;i<s.Length-1;i++)set.Add(s.Substring(i,2));return set; }
        static string FuzzyNormalize(string text){
            string s=Normalize(text);
            string[,] phrases={{"下面","下列"},{"以下","下列"},{"下述","下列"},{"哪一项","哪个"},{"哪种","哪个"},{"哪一","哪个"},{"哪个选项","哪个"},{"并非","不是"},{"不可以","不能"},{"不可","不能"},{"能够","能"},{"正确的是","正确"},{"错误的是","不正确"}};
            for(int i=0;i<phrases.GetLength(0);i++)s=s.Replace(phrases[i,0],phrases[i,1]);return s;
        }
        double CharacterSimilarity(string a,string b){
            var left=new HashSet<char>(a);var right=new HashSet<char>(b);double total=0,common=0;
            foreach(char c in left.Union(right)){List<int> ids;double weight=charIndex.TryGetValue(c,out ids)?1+Math.Log(1.0+questions.Count/(1.0+ids.Count)):2;bool x=left.Contains(c),y=right.Contains(c);total+=weight*((x?1:0)+(y?1:0));if(x&&y)common+=2*weight;}
            return total==0?0:common/total;
        }
        static double Dice(HashSet<string> a,HashSet<string> b) { return a.Count+b.Count==0?0:2.0*a.Count(t=>b.Contains(t))/(a.Count+b.Count); }
        static double Sequence(string a,string b) {
            if(a.Length==0||b.Length==0)return 0;
            if(a.Length>800)a=a.Substring(0,800);if(b.Length>800)b=b.Substring(0,800);
            var previous=new int[b.Length+1]; var current=new int[b.Length+1];
            for(int i=1;i<=a.Length;i++) { for(int j=1;j<=b.Length;j++)current[j]=a[i-1]==b[j-1]?previous[j-1]+1:Math.Max(previous[j],current[j-1]);var tmp=previous;previous=current;current=tmp; }
            return 2.0*previous[b.Length]/(a.Length+b.Length);
        }
        static bool Negative(string s) { return Regex.IsMatch(s,@"不是|不正确|不属于|不包括|不符合|不能|不应|错误|有误|除外|错误的|不属于|incorrect|except"); }
        public List<Hit> Search(string text,int limit) {
            if(string.IsNullOrWhiteSpace(text))return new List<Hit>();
            string stem=Regex.Split(text,@"(?m)^\s*[A-Ga-g]\s*[.、:：)）．]\s*")[0];
            stem=Regex.Replace(stem,@"(?m)^\s*(?:单选题|多选题|判断题|选择题|第\s*\d+\s*题|\d+\s*/\s*\d+)\s*$", "").Trim();
            string literal=Normalize(stem),query=FuzzyNormalize(stem);if(query.Length<4)return new List<Hit>();
            if(query.Length>1200)query=query.Substring(0,1200);
            var tokens=Tokens(query); var votes=new Dictionary<int,double>();
            foreach(var t in tokens) { List<int> ids; if(!index.TryGetValue(t,out ids))continue;double weight=1+Math.Log(1.0+questions.Count/(1.0+ids.Count));foreach(int id in ids){double value;votes.TryGetValue(id,out value);votes[id]=value+weight;} }
            // Individual characters recover candidates even when OCR breaks every adjacent pair.
            var characterVotes=new Dictionary<int,int>();
            foreach(char c in query.Distinct()){List<int> ids;if(!charIndex.TryGetValue(c,out ids))continue;foreach(int id in ids){int count;characterVotes.TryGetValue(id,out count);characterVotes[id]=count+1;}}
            var candidates=new HashSet<int>(votes.OrderByDescending(v=>v.Value).ThenBy(v=>v.Key).Take(100).Select(v=>v.Key));
            int distinct=query.Distinct().Count();
            foreach(int id in characterVotes.Where(v=>v.Value>=4 && v.Value/(double)Math.Min(distinct,questions[v.Key].Fuzzy.Distinct().Count())>=0.5).OrderByDescending(v=>v.Value/(double)Math.Max(distinct,questions[v.Key].Fuzzy.Distinct().Count())).ThenBy(v=>v.Key).Take(60).Select(v=>v.Key))candidates.Add(id);
            var hits=new List<Hit>();
            foreach(int id in candidates) {
                var q=questions[id];double d=Dice(tokens,q.Tokens),s=Sequence(query,q.Fuzzy),characters=CharacterSimilarity(query,q.Fuzzy);double score=0.45*d+0.25*s+0.30*characters;
                score=Math.Max(score,0.82*(0.75*characters+0.25*s));
                if(literal==q.Normal)score=1;
                else if(query==q.Fuzzy)score=0.98;
                else {
                    string meaningful=Regex.Replace(query,@"下列|哪个|选项|正确|的是|题目|选择|请问|什么|是否","");
                    if(meaningful.Length>=4 && q.Fuzzy.Contains(query))score=Math.Max(score,0.82+0.14*Math.Sqrt(query.Length/(double)q.Fuzzy.Length));
                    else if(q.Fuzzy.Length>=8&&query.Contains(q.Fuzzy))score=Math.Max(score,0.88*q.Fuzzy.Length/query.Length+0.1);
                }
                if(Negative(query)!=Negative(q.Fuzzy))score*=0.76;
                var queryNumbers=Regex.Matches(query,@"\d+").Cast<Match>().Select(m=>m.Value).ToArray();var bankNumbers=Regex.Matches(q.Fuzzy,@"\d+").Cast<Match>().Select(m=>m.Value).ToArray();
                if(queryNumbers.Length>0&&bankNumbers.Length>0&&!queryNumbers.SequenceEqual(bankNumbers))score*=0.72;
                string all=FuzzyNormalize(text); var optionTokens=Tokens(string.Concat(q.Options.Where(o=>!string.IsNullOrWhiteSpace(o)).Select(FuzzyNormalize)));
                if(optionTokens.Count>0 && all.Length>query.Length+4)score=score*0.92+0.08*optionTokens.Count(t=>all.Contains(t))/(double)optionTokens.Count;
                hits.Add(new Hit { Question=q,Score=Math.Min(100,score*100) });
            }
            return hits.OrderByDescending(h=>h.Score).ThenBy(h=>h.Question.Source).Take(limit).ToList();
        }
    }
}
