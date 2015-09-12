using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// TODO:    按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ExportJsonMenu();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意:  如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace ExportJson
{
    [ComVisible(true)]
    public class ExportJsonMenu : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ExportJsonMenu()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExportJson.ExportJsonMenu.xml");
        }

        #endregion

        #region 功能区回调
        //在此创建回调方法。有关添加回调方法的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }


        public void OnExportJson(Office.IRibbonControl control)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
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
                        string v = dataRang[i].Text;
                        Regex reg = new Regex(@"^[-]?[1-9]{1}\d*$|^[0]{1}$");
                        bool ismatch = reg.IsMatch(v);
                        if (ismatch)
                        {
                            stringBuilder.Append(Convert.ToInt64(v));
                        }
                        else
                        {
                            v = v.Replace('\r',' ');
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

            string path = Directory.GetCurrentDirectory() + @"/Json";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "json文件(*.json)|";
            saveFileDialog.ShowDialog();

            if(!saveFileDialog.FileName.EndsWith(".json"))
            {
                saveFileDialog.FileName += ".json";
            }

            FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.Write(json);
            sw.Close();
        }

        List<Excel.Range> GetLine(Excel.Worksheet sheet, int lineNum)
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

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion


    }
}
