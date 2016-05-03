using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using fastJSON;

namespace ExportJsonPlugin
{
    public partial class ExportJson
    {
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
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "json文件(*.json)|";
            string fileName = Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            saveFileDialog.FileName = fileName;
            saveFileDialog.ShowDialog();

            string jsonFileName = "";
            if (!saveFileDialog.FileName.EndsWith(".json"))
            {
                jsonFileName = saveFileDialog.FileName + ".json";
            }

            FileStream fs = new FileStream(jsonFileName, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(json);
            sw.Close();

            
            //             ExportCShape ex = new ExportCShape(saveFileDialog.FileName);
            //             ex.AddField("描述...", FieldType.Int, "Name");
            //             ex.Finish();
            
            MessageBox.Show("导出完成");
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

            List<Excel.Range> keyRang = GetLine(activeWorksheet, 1);
            if (keyRang.Count == 0)
            {
                return;
            }
            foreach (Excel.Range cell in keyRang)
            {
                _Keys.Add(cell.Text);
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
                List<Excel.Range> dataRang = GetLine(activeWorksheet, ++index);
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
                        StringBuilder v1 = new StringBuilder(dataRang[i].Text);
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
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            if (activeWorksheet == null)
            {
                MessageBox.Show("请启动编辑模式");
                return;
            }
            List<string> _Keys = new List<string>();
            //字段类型
            List<Excel.Range> typeRang = GetLine(activeWorksheet, 2);
            //字段名称
            List<Excel.Range> keyRang = GetLine(activeWorksheet, 3);
            if (typeRang.Count != keyRang.Count || typeRang.Count == 0 || keyRang.Count == 0)
            {
                MessageBox.Show("字段和类型个数不匹配");
                return;
            }
            Dictionary<string, string> _FieldsDic = new Dictionary<string, string>();
            for (int i = 0; i < keyRang.Count; ++i)
            {
                Excel.Range cell = keyRang[i];
                _FieldsDic.Add(cell.Text, typeRang[i].Text);
                _Keys.Add(cell.Text);
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
                List<Excel.Range> dataRang = GetLine(activeWorksheet, ++index);
                if (dataRang.Count == 0)
                {
                    break;
                }

                stringBuilder.Append("{");

                for (int i = 0; i < _Keys.Count; ++i)
                {
                    stringBuilder.Append("\"" + ((string)keyRang[i].Text).Trim() + "\":");
                    string fieldType = ((string)typeRang[i].Text).Trim();
                    if (i < dataRang.Count)
                    {
                        StringBuilder v1 = new StringBuilder(dataRang[i].Text);
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
                                    stringBuilder.Append(Convert.ToDouble(v));
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
            this.Save(json);
        }


        private List<Excel.Range> GetLine(Excel.Worksheet sheet, int lineNum)
        {
            char a = 'A';
            char z = 'Z';

            int col = 1000;

            string colStr = "";
            List<Excel.Range> _Cells = new List<Excel.Range>();

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
                if (string.IsNullOrEmpty(line.Text))
                {
                    break;
                }
                _Cells.Add(line);
            }
            return _Cells;
        }
    }
}