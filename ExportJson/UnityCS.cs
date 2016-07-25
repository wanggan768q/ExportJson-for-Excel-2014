using ExportJson.Properties;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportJsonPlugin
{
    class UnityCS
    {
        public const string T1 = "\t";
        public const string T2 = "\t\t";
        public const string T3 = "\t\t\t";
        public const string T4 = "\t\t\t\t";

        Dictionary<string, StringBuilder> templateDir = new Dictionary<string, System.Text.StringBuilder>()
        {
            {"$Template$",new StringBuilder() },
            {"$FieldDefine$",new StringBuilder() },
            {"$ColCount$",new StringBuilder() },
            {"$CheckColName$",new StringBuilder() },
            {"$ReadBinColValue$",new StringBuilder() },
            {"$ReadCsvColValue$",new StringBuilder() },
        };

        string name = "";
        StringBuilder sb = new StringBuilder();
        public void Export(string fileName,List<string> type, List<string> key,List<string> des)
        {
            string[] titles = fileName.Split('_');
            name = titles[0];

            templateDir["$Template$"].Append(name);

            AddField(templateDir["$FieldDefine$"], type, key,des);

            templateDir["$ColCount$"].Append(key.Count);

            CheckColName(templateDir["$CheckColName$"],key,name);

            ReadBinColValue(templateDir["$ReadBinColValue$"], key);

            ReadCsvColValue(templateDir["$ReadCsvColValue$"], type, key);

            string text = Resources.ConfigTemplate;

            foreach(var dir in templateDir)
            {
                text = text.Replace(dir.Key, dir.Value.ToString());
            }

            string path = Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName) + Path.DirectorySeparatorChar + "Out";
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            path += Path.DirectorySeparatorChar + fileName + ".cs";
            StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8);
            sw.Write(text);
            sw.Flush();
            sw.Close();

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
                        sb.AppendFormat(T3 + E("member.{0} = Convert.ToDouble(vecLine[{1}]);"), key[i], i);
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

        void ReadBinColValue(StringBuilder sb, List<string> key)
        {
            for (int i = 0; i < key.Count; ++i)
            {
                sb.AppendFormat(T3 + E("readPos += HS_ByteRead.ReadInt32Variant(binContent, readPos, out member.{0} );"), i, key[i]);
            }
        }


        string E(string s="")
        {
            return s + "\r\n";
        }

        void AddTitle(StringBuilder sb,string title)
        {
            sb.AppendLine("/// <summary>");
            sb.AppendFormat(E("/// {0}"), title);
            sb.AppendLine("/// </summary>");
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
                AddTitle(sb,des[i]);
                sb.AppendFormat(T1 + E("public {0} {1};            " + T1), t, k);
            }
        }
        
    }
}
