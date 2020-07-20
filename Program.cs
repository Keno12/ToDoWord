using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ToDoWord
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //01 check file dir
            Console.WriteLine("请输入文件夹路径:");
            //string filePath = CheckPath();
            string filePath = @"D:\迅雷下载";
            Console.WriteLine("-----------------------");

            //02 trasn doc to docx
            Console.WriteLine("将所有的.doc文件换为.docx");
            DealDocFiles(filePath);

            //03 deal with docx file
            Console.WriteLine("开始处理.docx文件");
            GetFiles(filePath);

            //04 end out msg
            Console.WriteLine();
            Console.WriteLine("点击任意键退出");
            Console.ReadKey();
        }

        private static string CheckPath()
        {
            string path = Console.ReadLine();
            if (Directory.Exists(path))
            {
                return path;
            }
            else
            {
                Console.WriteLine("文件路径不存在请重新输入,请输入文件路径:");
                return CheckPath();
            }
        }

        //get all files
        private static void GetFiles(string filePath)
        {
            string[] files = Directory.GetFiles(filePath, @"*.docx");
            Console.WriteLine("检测到共:{0}个.docx文件", files.Count());
            int i = 0;
            foreach (string file in files)
            {
                i++;
                Console.WriteLine("正在处理第:{0}个.docx文件", i);
                DealWithWord(file);
                //FindReplaceText(file);
            }
            Console.WriteLine("-----------------------");
            Console.WriteLine();
            Console.WriteLine("所有.docx文件处理完成!");
        }

        //deal with doc
        private static void DealDocFiles(string filePath)
        {
            string[] docfiles = Directory.GetFiles(filePath, "*.doc");
            List<string> relDocFils = new List<string>();
            foreach (var docfile in docfiles)
            {
                if (docfile.Contains(".docx"))
                {
                    continue;
                }
                relDocFils.Add(docfile);
            }
            string[] docxfiles = Directory.GetFiles(filePath, "*.docx");
            Console.WriteLine("检测到共:{0}个.doc文件", relDocFils.Count());

            int i = 0;
            StringBuilder sb = new StringBuilder();
            foreach (string file in relDocFils)
            {
                string fileName = file + "x";
                if (docxfiles.Contains(fileName))
                {
                    sb.AppendLine(file);
                    continue;
                }
                i++;
                Console.WriteLine("正在处理第:{0}个.doc文件", i);
                TranWordDocToDocx(file);
            }
            Console.WriteLine();
            Console.WriteLine("所有.doc文件处理完成,其中这些.doc文件已存在对应的.docx文件未做处理");
            Console.WriteLine("-----------------------");
            Console.Write(sb);
            Console.WriteLine("-----------------------");
            Console.WriteLine();
        }

        private static void DealWithWord(string path)
        {
            //using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            //{
            //    MainDocumentPart mainPart = doc.MainDocumentPart;
            //    Hyperlink hLink = mainPart.Document.Body.Descendants<Hyperlink>().FirstOrDefault();
            //    if (hLink != null)
            //    {
            //        // get hyperlink's relation Id (where path stores)
            //        string relationId = hLink.Id;
            //        if (relationId != string.Empty)
            //        {
            //            // get current relation
            //            HyperlinkRelationship hr = mainPart.HyperlinkRelationships.Where(a => a.Id == relationId).FirstOrDefault();
            //            if (hr != null)
            //            // remove current relation
            //            { mainPart.DeleteReferenceRelationship(hr); }
            //            //add new relation with same Id , but new path
            //            mainPart.AddHyperlinkRelationship(new System.Uri(@"D:\work\DOCS\new\My.docx", System.UriKind.Absolute), true, relationId);
            //        }
            //    }
            //    // apply changes
            //    doc.Close();
            //}

            //02  code demo
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                foreach (OpenXmlElement paragraph in body.Elements())
                {
                    foreach (OpenXmlElement item in paragraph.Elements<OpenXmlElement>())
                    {
                        if (item.InnerText.Contains("HYPERLINK") || item.InnerText.Contains("hyperlink"))
                        {
                            Regex regexText = new Regex("[^\x00-\xff]");
                            if (regexText.IsMatch(item.InnerText))
                            {
                                //01  copy this text 
                                string thisInnerText = item.InnerText;
                                //02  remove link
                                Regex regexLink = new Regex("^((https|http|ftp|rtsp|mms)?:\\/\\/)[^\\s]+");
                                foreach (Match item1 in regexLink.Matches(thisInnerText))
                                {

                                }
                                //03  insert text at same place
                            }
                            else
                            {
                                item.Remove();
                            }

                        }
                    }
                }
                doc.Save();
            }
        }

        //find the text in the  doc
        private static void FindReplaceText(string path)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("<w:hyperlink[\\w\\W]*history=\\\"(\\d|-\\d)\\\">");
                var matches = regexText.Matches(docText);
                foreach (Match item in matches)
                {
                }
                docText = regexText.Replace(docText, "121212");

                using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private static void TranWordDocToDocx(string file)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object oMissing = System.Reflection.Missing.Value;

            Object ConfirmConversions = false;
            Object ReadOnly = false;
            Object AddToRecentFiles = false;

            Object PasswordDocument = "";
            Object PasswordTemplate = System.Type.Missing;
            Object Revert = System.Type.Missing;
            Object WritePasswordDocument = System.Type.Missing;
            Object WritePasswordTemplate = System.Type.Missing;
            Object Format = System.Type.Missing;
            Object Encoding = System.Type.Missing;
            Object Visible = System.Type.Missing;
            Object OpenAndRepair = System.Type.Missing;
            Object DocumentDirection = System.Type.Missing;
            Object NoEncodingDialog = System.Type.Missing;
            Object XMLTransform = System.Type.Missing;

            Object fileName = file;
            doc = word.Documents.Open(ref fileName, ref ConfirmConversions,
            ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate,
            ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format,
            ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection,
            ref NoEncodingDialog, ref XMLTransform);

            object path = file.Substring(0, file.Length - 4);
            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
            doc.SaveAs(ref path, ref format, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Close(ref oMissing, ref oMissing, ref oMissing);
            word.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
    }
}