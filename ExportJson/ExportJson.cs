using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using fastJSON;
using ExportJson.Properties;
using System.Linq;

namespace ExportJsonPlugin
{
    public partial class ExportJson
    {
        private string WorkbookFullName
        {
            get
            {
                return Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            }
        }
        private void ExportJson_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void Save(string json)
        {
            try
            {
                JSON.Instance.Parse(json);
            }
            catch (System.Exception _ex)
            {
                MessageBox.Show("数据异常,请检查");
                return;
            }
//             SaveFileDialog saveFileDialog = new SaveFileDialog();
//             saveFileDialog.Filter = "json文件(*.json)|";
//             string fileName = Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
//             string[] titles = fileName.Split('_');
//             string name = titles[0];
//             saveFileDialog.InitialDirectory = Path.GetDirectoryName(fileName);
//             saveFileDialog.FileName = name;
//             saveFileDialog.ShowDialog();

            string fileName = Path.GetFileNameWithoutExtension(WorkbookFullName);
            string[] titles = fileName.Split('_');
            string name = titles[0];

            string path = Path.GetDirectoryName(WorkbookFullName) + Path.DirectorySeparatorChar + "Out" + Path.DirectorySeparatorChar + "Json";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string jsonFileName = "";
            if (!fileName.EndsWith(".json"))
            {
                //jsonFileName = saveFileDialog.FileName + ".json";
                jsonFileName = path + Path.DirectorySeparatorChar + name + ".json";
            }

            if (jsonFileName.Contains(":"))
            {
                FileStream fs = new FileStream(jsonFileName, FileMode.CreateNew | FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(json);
                sw.Close();
                MessageBox.Show("导出完成");
            }
        }

        /// <summary>
        /// 把所有数据导出
        /// 1.字段
        /// 2.数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportJsonOfNormal(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            if(activeWorksheet == null)
            {
                MessageBox.Show("请启动编辑模式");
                return;
            }
            List<string> _Keys = new List<string>();

            List<string> keyRang = GetLine(activeWorksheet, 1);
            if (keyRang.Count == 0)
            {
                return;
            }
            foreach (string cell in keyRang)
            {
                _Keys.Add(cell);
            }
            if (_Keys.Count == 0)
            {
                return;
            }

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("[");

            int keyCount = keyRang.Count;
            int index = 1;
            while (true)
            {
                List<string> dataRang = GetLine(activeWorksheet, ++index);
                if (dataRang.Count == 0)
                {
                    break;
                }
                stringBuilder.Append("{");

                for (int i = 0; i < _Keys.Count; ++i)
                {
                    stringBuilder.Append("\"" + _Keys[i] + "\":");
                    if (i < dataRang.Count)
                    {
                        StringBuilder v1 = new StringBuilder(dataRang[i]);
                        //string v = dataRang[i].Text;
                        string v = v1.ToString().TrimEnd();
                        bool isInteger = Regex.IsMatch(v, @"^[-]?[1-9]{1}\d*$|^[0]{1}$");
                        bool isDecimal = Regex.IsMatch(v, @"^(-?\d+)(\.\d+)?$");
                        if (isInteger)
                        {
                            stringBuilder.Append(Convert.ToInt64(v));
                        }
                        else if (isDecimal)
                        {
                            stringBuilder.Append(Convert.ToDouble(v));
                        }
                        else
                        {
                            //v = v.Replace('\r', ' ');
                            v = v.Replace('\n', ' ');
                            stringBuilder.Append("\"" + v + "\"");
                        }
                    }
                    else
                    {
                        stringBuilder.Append("");
                    }
                    if (i != _Keys.Count - 1)
                    {
                        stringBuilder.Append(",");
                    }
                }
                stringBuilder.Append("},");
            }
            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            stringBuilder.Append("]");
            string json = stringBuilder.ToString();
            /*
            byte[] unicodeBuf = Encoding.Unicode.GetBytes(json);
            byte[] utfBuf = Encoding.Convert(Encoding.Unicode, Encoding.UTF8, unicodeBuf);
            json = Encoding.UTF8.GetString(utfBuf);
             * */
            //验证JSON
            /*
            string utfJson = GB2312ToUTF8(json);

            JsonData dataSrc = JsonMapper.ToObject(utfJson);
            int a = 0;
            */

            this.Save(json);
        }


        //字段类型
        List<string> typeRang = new List<string>();
        //字段名称
        List<string> keyRang = new List<string>();

        /// <summary>
        /// 根据类型导出数据
        /// 1.描述
        /// 2.类型  I->int F->float S->String B->bool
        /// 3.字段
        /// 4.数据  一级分隔符| 二级分隔符 _
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportJsonOfType(object sender, RibbonControlEventArgs e)
        {
            typeRang.Clear();
            keyRang.Clear();

            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            string[] fileName = activeWorksheet.Application.Caption.Split('.');

            List<string> Keys = new List<string>();

            List<string> Des = new List<string>();
            Des = GetLine(activeWorksheet, 1);

            keyRang = GetLine(activeWorksheet, 3);

            typeRang = GetLine(activeWorksheet, 2);

            

            if (typeRang.Count != keyRang.Count || typeRang.Count == 0 || keyRang.Count == 0)
            {
                MessageBox.Show("字段和类型个数不匹配");
                return;
            }
            Dictionary<string, string> _FieldsDic = new Dictionary<string, string>();
            for (int i = 0; i < keyRang.Count; ++i)
            {
                string cell = keyRang[i];
                _FieldsDic.Add(cell, typeRang[i]);
                Keys.Add(cell);
            }

            if (_FieldsDic.Count == 0)
            {
                return;
            }
            
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("[");
            int index = 3;
            while (true)
            {
                List<string> dataRang = GetLine(activeWorksheet, ++index,false);
                if (dataRang.Count == 0)
                {
                    break;
                }

                stringBuilder.Append("{");

                for (int i = 0; i < Keys.Count; ++i)
                {
                    stringBuilder.Append("\"" + ((string)keyRang[i]).Trim() + "\":");
                    string fieldType = ((string)typeRang[i]).Trim();
                    if (i < dataRang.Count)
                    {
                        StringBuilder v1 = new StringBuilder(dataRang[i]);
                        //string v = dataRang[i].Text;
                        string v = v1.ToString().TrimEnd();

                        switch (fieldType)
                        {
                            case "I":
                                {
                                    stringBuilder.Append(Convert.ToInt64(v));
                                }
                                break;
                            case "F":
                                {
                                    stringBuilder.AppendFormat("{0:F}",Convert.ToDouble(v));
                                }
                                break;
                            case "S":
                                {
                                    v = v.Replace('\r', ' ');
                                    v = v.Replace('\n', ' ');
                                    stringBuilder.Append("\"" + v + "\"");
                                }
                                break;
                            case "B":
                                {
                                    v = v.Trim();
                                    v = v.ToLower();
                                    if (v.Equals("0") || v.Equals(bool.FalseString.ToLower()))
                                    {
                                        stringBuilder.Append(bool.FalseString.ToLower());
                                    }
                                    else if (v.Equals("1") || v.Equals(bool.TrueString.ToLower()))
                                    {
                                        stringBuilder.Append(bool.TrueString.ToLower());
                                    }
                                    else
                                    {
                                        MessageBox.Show("错误的类型: [ " + fieldType + " ]");
                                    }
                                }
                                break;
                            default:
                                {
                                    MessageBox.Show("错误的类型: [ " + fieldType + " ]");
                                }
                                break;
                        }
                    }
                    else
                    {
                        stringBuilder.Append("");
                    }
                    if (i != Keys.Count - 1)
                    {
                        stringBuilder.Append(",");
                    }
                }
                stringBuilder.Append("},");
            }
            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            stringBuilder.Append("]");
            string json = stringBuilder.ToString();
            this.Save(json);

            UnityCS cs = new UnityCS();
            cs.Export(fileName[0], typeRang, keyRang,Des);
        }


        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="lineNum"></param>
        /// <param name="assert">是否判null</param>
        /// <returns></returns>
        private List<string> GetLine(Excel.Worksheet sheet, int lineNum,bool assert = true)
        {
            char a = 'A';
            char z = 'Z';

            int col = 1000;

            string colStr = "";
            List<string> _Cells = new List<string>();

            for (int c = 0; c < col; ++c)
            {
                string rStr = string.Format("{0}{1}", colStr + a, lineNum);
                a++;
                if (a > z)
                {
                    a = 'A';
                    colStr += a;
                }
                Excel.Range line = sheet.get_Range(rStr);
                string tempV = line.Text;
                if (string.IsNullOrEmpty(tempV))
                {
                    if(assert || c == typeRang.Count)
                    {
                        break;
                    }
                    if(c < typeRang.Count)
                    {
                        string type = ((string)typeRang[c]).Trim();
                        switch (type)
                        {
                            case "I":
                            case "F":
                            case "B":
                                {
                                    tempV = "0";
                                }
                                break;
                            case "S":
                                {
                                    tempV = "";
                                }
                                break;
                            default:
                                {
                                    MessageBox.Show("错误的类型: [ " + type + " ]");
                                }
                                break;
                        }
                    }
                }
                _Cells.Add(tempV);
            }

            int index = 0;
            List<string> tempDataRang = _Cells.FindAll((string s) =>
            {
                if (string.IsNullOrEmpty(s) || s.Equals("0"))
                {
                    if (++index == _Cells.Count)
                    {
                        _Cells.Clear();
                    }
                    return true;
                }
                
                return false;
            });
            return _Cells;
        }
    }
}