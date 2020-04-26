using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks; 
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc; 
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Reflection;

namespace OfficeService.Controllers
{  
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {
        // GET: api/Word
        [HttpGet]
        public string Get()
        {
            return "it works with office 2013 installed.";
        }
         
        // POST: api/word/replace 
        [HttpPost("[action]")]
        public IActionResult replace([FromForm]Doc doc) {
            try { 
                dynamic jsons = JsonConvert.DeserializeObject(doc.json);

                Dictionary<string, string> dicValues = new Dictionary<string, string>();

                foreach (var item in jsons) {
                    dicValues.Add((string)item.Path, (string)item.Value);
                }


                string oldPath = AppContext.BaseDirectory + "old.docx";
                System.IO.File.WriteAllBytes(oldPath, Convert.FromBase64String(doc.base64));

                string newPath = AppContext.BaseDirectory + "new.docx";

                WordReplace(oldPath, newPath, dicValues);

                byte[] newBytes = System.IO.File.ReadAllBytes(newPath);

                try
                {
                    System.IO.File.Delete(oldPath);
                    System.IO.File.Delete(newPath);
                }
                catch { }

                return Ok(new { docBase64 = Convert.ToBase64String( newBytes ) });
            }
            catch (Exception ex) {
                return Ok(new { exception = ex.Message });
            }
        }

        public class Doc { 
            public string base64 { get; set; }
            public string json { get; set; }
        }
 

        public static void WordReplace(string oldWordPath, string newWordPath, Dictionary<string, string> dicValues ) {

            Object Nothing = Missing.Value; //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
            //object format = WdSaveFormat.wdFormatDocumentDefault;
            //object unite = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            object newDoc = newWordPath;
            Application wordApp;//Word应用程序变量初始化
            Document wordDoc;  

            wordApp = new Application();//创建word应用程序

            object fileName = (oldWordPath);//模板文件

            wordDoc = wordApp.Documents.Open(ref fileName,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            object replace = WdReplace.wdReplaceAll;

            wordApp.Selection.Find.Replacement.ClearFormatting();
            wordApp.Selection.Find.MatchWholeWord = true;
            wordApp.Selection.Find.ClearFormatting();

            foreach (var item in dicValues) {
                object FindText = item.Key;
                object Replacement = item.Value;

                if (Replacement.ToString().Length > 110)
                {
                    FindAndReplaceLong(wordApp, FindText, Replacement);
                }else { 
                    FindAndReplace(wordApp, FindText, Replacement); 
                }
            }
            
            wordDoc.SaveAs(newDoc,
            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
            //关闭wordDoc文档
            wordApp.Documents.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

        }


        public static void FindAndReplaceLong(Application wordApp, object findText, object replaceText)
        {
            int len = replaceText.ToString().Length; //要替换的文字长度
            int cnt = len / 110; //不超过220个字
            string newstr;
            object newStrs;
            if (len < 110) //小于220字直接替换
            {
                FindAndReplace(wordApp, findText, replaceText);
            }
            else
            {
                for (int i = 0; i <= cnt; i++)
                {
                    if (i != cnt)
                        newstr = replaceText.ToString().Substring(i * 110, 110) + findText; //新的替换字符串
                    else
                        newstr = replaceText.ToString().Substring(i * 110, len - i * 110); //最后一段需要替换的文字
                    newStrs = (object)newstr;
                    FindAndReplace(wordApp, findText, newStrs); //进行替换
                }
            }
        }

        public static void FindAndReplace(Application wordApp, object findText, object replaceText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
            ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceText,
            ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }



    }
}
