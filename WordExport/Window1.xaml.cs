using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WordExport
{
    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    public partial class Window1 : System.Windows.Window
    {
        public Window1()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
                string wordTemplatePath = System.Windows.Forms.Application.StartupPath + @"\Word模板.docx";
                if (File.Exists(wordTemplatePath))
                {
                    System.Windows.Forms.FolderBrowserDialog dirDialog = new System.Windows.Forms.FolderBrowserDialog();
                    dirDialog.ShowDialog();
                    if (dirDialog.SelectedPath != string.Empty)
                    {
                        string newFileName = dirDialog.SelectedPath + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx";
                        
                        Dictionary<string, string> wordLableList = new Dictionary<string, string>();
                        wordLableList.Add("年", "2021");
                        wordLableList.Add("月", "9");
                        wordLableList.Add("日", "18");
                        wordLableList.Add("星期", "六");
                        wordLableList.Add("标题", "Word导出数据");
                        wordLableList.Add("内容", "我是内容——Kiba518");

                        Export(wordTemplatePath, newFileName, wordLableList);
                        MessageBox.Show("导出成功!");
                    }
                    else
                    {
                        MessageBox.Show("请选择导出位置");
                    } 
                }
                else
                { 
                    MessageBox.Show("Word模板文件不存在!"); 
                } 
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString());
                return;
            }
        }
        public static void Export(string wordTemplatePath, string newFileName, Dictionary<string, string> wordLableList)
        {  
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string TemplateFile = wordTemplatePath;
            File.Copy(TemplateFile, newFileName);
            _Document doc = new Document();
            object obj_NewFileName = newFileName;
            object obj_Visible = false;
            object obj_ReadOnly = false;
            object obj_missing = System.Reflection.Missing.Value;
           
            doc = app.Documents.Open(ref obj_NewFileName, ref obj_missing, ref obj_ReadOnly, ref obj_missing,
                ref obj_missing, ref obj_missing, ref obj_missing, ref obj_missing,
                ref obj_missing, ref obj_missing, ref obj_missing, ref obj_Visible,
                ref obj_missing, ref obj_missing, ref obj_missing,
                ref obj_missing);
            doc.Activate();

            if (wordLableList.Count > 0)
            {
                object what = WdGoToItem.wdGoToBookmark; 
                foreach (var item in wordLableList)
                {
                    object lableName = item.Key;
                    if (doc.Bookmarks.Exists(item.Key))
                    {
                        doc.ActiveWindow.Selection.GoTo(ref what, ref obj_missing, ref obj_missing, ref lableName);//光标移动书签的位置
                        doc.ActiveWindow.Selection.TypeText(item.Value);//在书签处插入的内容 
                        doc.ActiveWindow.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//设置插入内容的Alignment
                    }  
                }
            }

            object obj_IsSave = true;
            doc.Close(ref obj_IsSave, ref obj_missing, ref obj_missing);

        }
    }
}
