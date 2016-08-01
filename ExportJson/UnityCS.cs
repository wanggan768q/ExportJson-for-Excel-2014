using ExportJson.Properties;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExportJsonPlugin
{
    class UnityCS
    {
        public const string T1 = "\t";
        public const string T2 = "\t\t";
        public const string T3 = "\t\t\t";
        public const string T4 = "\t\t\t\t";

        private string WorkbookFullName
        {
            get
            {
                return Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            }
        }

        private string CSPath
        {
            get
            {
                return Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName) + Path.DirectorySeparatorChar + "Out" + Path.DirectorySeparatorChar + "config";
            }
        }


        Dictionary<string, StringBuilder> templateDir = new Dictionary<string, System.Text.StringBuilder>()
        {
            {"$Template$",new StringBuilder() },
            {"$FieldDefine$",new StringBuilder() },
            {"$ColCount$",new StringBuilder() },
            {"$CheckColName$",new StringBuilder() },
            {"$ReadBinColValue$",new StringBuilder() },
            {"$ReadCsvColValue$",new StringBuilder() },
            {"$InitPrimaryField$",new StringBuilder() },
            {"$PrimaryKey$",new StringBuilder() },
            {"$ReadJsonColValue$",new StringBuilder() },
        };

        string name = "";
        StringBuilder sb = new StringBuilder();
        public void Export(string fileName, List<string> type, List<string> key, List<string> des)
        {
            string[] titles = fileName.Split('_');
            name = titles[0];

            templateDir["$Template$"].Append(name);
            templateDir["$PrimaryKey$"].Append(key[0]);

            AddField(templateDir["$FieldDefine$"], type, key,des);

            templateDir["$ColCount$"].Append(key.Count);

            CheckColName(templateDir["$CheckColName$"],key,name);

            ReadBinColValue(templateDir["$ReadBinColValue$"], type, key);

            ReadCsvColValue(templateDir["$ReadCsvColValue$"], type, key);

            InitPrimaryField(templateDir["$InitPrimaryField$"], templateDir["$PrimaryKey$"].ToString());

            ReadJsonColValue(templateDir["$ReadJsonColValue$"],type,key);

            string text = Resources.ConfigTemplate;

            foreach(var dir in templateDir)
            {
                text = text.Replace(dir.Key, dir.Value.ToString());
            }

            
            if(!Directory.Exists(CSPath))
            {
                Directory.CreateDirectory(CSPath);
            }
            string path = CSPath+ Path.DirectorySeparatorChar + name + ".cs";
            StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
            sw.Write(text);
            sw.Flush();
            sw.Close();

            GenerateConfigLoad();
        }

        private void GenerateConfigLoad()
        {
            Dictionary<string, StringBuilder> dir = new Dictionary<string, System.Text.StringBuilder>()
            {
                {"$loadConfItem$",new StringBuilder() },
                {"$fileCount$",new StringBuilder() }
            };
            int fileCount = 0;
            string[] files = Directory.GetFiles(Path.GetDirectoryName(WorkbookFullName));
            foreach (string s in files)
            {
                string file = Path.GetFileName(s).Split('_')[0];
                string suffix = ".json";
                if (!file.Contains('~'))
                {
                    fileCount++;
                    StringBuilder sb = dir["$loadConfItem$"];
                    sb.AppendFormat(T2 + E("yield return StartCoroutine(LoadData(\"{0}" + suffix + "\"));"), file);
                    sb.AppendFormat(T2 + E("{0}Table.Instance.LoadJson(textContent);"), file);
                    sb.AppendFormat(T2 + E("Progress({0});"), fileCount);
                }
            }
            dir["$fileCount$"].Append(fileCount);

            string text = Resources.ConfigLoadTemplate;
            foreach (var d in dir)
            {
                text = text.Replace(d.Key, d.Value.ToString());
            }
            string path = CSPath + Path.DirectorySeparatorChar + "ConfigLoad.cs";
            StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
            sw.Write(text);
            sw.Flush();
            sw.Close();
        }

        void ReadJsonColValue(StringBuilder sb, List<string> type, List<string> key)
        {
            for (int i = 0; i < key.Count; ++i)
            {
                switch (type[i])
                {
                    case "I":
                        sb.AppendFormat(T3 + E("member.{0} = (int)jd[\"{1}\"];"), key[i], key[i]);
                        break;
                    case "F":
                        sb.AppendFormat(T3 + E("member.{0} = (float)((double)jd[\"{1}\"]);"), key[i], key[i]);
                        break;
                    case "B":
                        sb.AppendFormat(T3 + E("member.{0} = (bool)jd[\"{1}\"];"), key[i], key[i]);
                        break;
                    case "S":
                        sb.AppendFormat(T3 + E("member.{0} = (string)jd[\"{1}\"];"), key[i], key[i]);
                        break;
                }
            }
        }

        void InitPrimaryField(StringBuilder sb, string key)
        {
            sb.AppendLine(T2 + key + " = 0;");
            sb.AppendLine(T2 + "IsValidate = false;");
        }

        void ReadCsvColValue(StringBuilder sb, List<string> type, List<string> key)
        {
            for (int i=0;i<key.Count;++i)
            {
                switch(type[i])
                {
                    case "I":
                        sb.AppendFormat(T3 + E("member.{0} = Convert.ToInt32(vecLine[{1}]);"), key[i], i);
                        break;
                    case "F":
                        sb.AppendFormat(T3 + E("member.{0} = (float)Convert.ToDouble(vecLine[{1}]);"), key[i], i);
                        break;
                    case "B":
                        sb.AppendFormat(T3 + E("member.{0 }= Convert.ToBoolean(vecLine[{1}]);"), key[i], i);
                        break;
                    case "S":
                        sb.AppendFormat(T3 + E("member.{0} = vecLine[{1}];"), key[i], i);
                        break;
                }
            }
        }

        
        void CheckColName(StringBuilder sb,List<string> key,string name)
        {
            for (int i = 0; i < key.Count; ++i)
            {
                sb.AppendFormat(T2 + E("if(vecLine[{0}]!=\"{1}\") {{ Debug.Log(\"{2}.json中字段[{3}]位置不对应\"); return false; }}"),
                    i, key[i], name, key[i], name);
            }
            sb.AppendLine();
        }

        void ReadBinColValue(StringBuilder sb,List<string> type,List<string> key)
        {
            for (int i = 0; i < key.Count; ++i)
            {
                switch (type[i])
                {
                    case "I":
                        sb.AppendFormat(T3 + E("readPos += HS_ByteRead.ReadInt32Variant(binContent, readPos, out member.{0} );"), key[i]);
                        break;
                    case "F":
                        sb.AppendFormat(T3 + E("readPos += HS_ByteRead.ReadFloat(binContent, readPos, out member.{0} );"), key[i]);
                        break;
                    case "B":
                        sb.AppendFormat(T3 + E("readPos += HS_ByteRead.ReadBool(binContent, readPos, out member.{0} );"), key[i]);
                        break;
                    case "S":
                        sb.AppendFormat(T3 + E("readPos += HS_ByteRead.ReadString(binContent, readPos, out member.{0} );"), key[i]);
                        break;
                }
                
            }
        }


        string E(string s="")
        {
            return s + "\r\n";
        }

        void AddTitle(StringBuilder sb,string title)
        {
            sb.AppendLine(T1 + "/// <summary>");
            sb.AppendFormat(E(T1 + "/// {0}"), title);
            sb.AppendLine(T1 + "/// </summary>");
        }

        /// <summary>
        /// 添加字段
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="type"></param>
        /// <param name="key"></param>
        /// <param name="des"></param>
        void AddField(StringBuilder sb,List<string> type, List<string> key, List<string> des)
        {
            for(int i=0;i<type.Count;++i)
            {
                string t = type[i];
                string k = key[i];
                switch (t)
                {
                    case "I":
                        t = "int";
                        break;
                    case "F":
                        t = "float";
                        break;
                    case "B":
                        t = "bool";
                        break;
                    case "S":
                        t = "string";
                        break;
                }
                AddTitle(sb,des[i].Replace("\n","\t"));
                sb.AppendFormat(T1 + E("public {0} {1};"), t, k);
                sb.AppendLine();
            }
        }
        
    }
}
