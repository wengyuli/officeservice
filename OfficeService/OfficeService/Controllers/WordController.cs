using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks; 
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text;
using Newtonsoft.Json;

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
            return "it works.";
        }
         
        // POST: api/word/replace 
        [HttpPost("[action]")]
        public IActionResult replace([FromForm]Doc doc) {
            try {
                // {"str1" : "newStr1", "str2" : "newStr2"}; 
                dynamic jsons = JsonConvert.DeserializeObject(doc.json);

                Dictionary<string, string> dicValues = new Dictionary<string, string>();

                foreach (var item in jsons) {
                    dicValues.Add((string)item.Path, (string)item.Value);
                }

                var base64 = ReplaceContent(doc.base64, dicValues);

                return Ok(new { docBase64 = base64 });
            }
            catch (Exception ex) {
                return Ok(new { exception = ex.Message });
            }
        }

        public class Doc { 
            public string base64 { get; set; }
            public string json { get; set; }
        }

        public static string ReplaceContent(string wordBase64, Dictionary<string, string> replaysDictionary)
        {
            try {
                IWordDocument document = new WordDocument();
                var wordBytes = Encoding.UTF8.GetBytes(wordBase64);
                var fileMemoryStream = new MemoryStream(wordBytes);
                document.Open(fileMemoryStream, FormatType.Doc);
                foreach (var rd in replaysDictionary)
                {
                    if (string.IsNullOrEmpty(document.GetText())) continue;

                    document.Replace(rd.Key, rd.Value, false, false);
                    while (document.GetText().IndexOf(rd.Key) != -1)
                        document.Replace(rd.Key, rd.Value, false, false);
                }
                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Doc);

                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                stream.Seek(0, SeekOrigin.Begin);

                document.Close();
                stream.Position = 0;

                return Encoding.UTF8.GetString(bytes);
            }
            catch (Exception ex) {
                return Encoding.UTF8.GetString( Encoding.UTF8.GetBytes( ex.Message) );
            }
            
        }



    }
}
