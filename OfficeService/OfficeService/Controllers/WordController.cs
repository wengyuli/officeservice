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

        // GET: api/Word/5
        [HttpGet("{id}", Name = "Get")]
        public string Get(int id)
        {
            return "word api";
        }

        // POST: api/Word
        [HttpPost]
        public void Post([FromBody] string value)
        {

        } 

        // POST: api/word/replace 
        [HttpPost("replace")]
        public string replace([FromBody]string request) {
            return "000";
        }
         
        public static string ReplaceDocContent(string wordBase64, Dictionary<string, string> replaysDictionary)
        {
            IWordDocument document = new WordDocument();
            var wordBytes = Encoding.UTF8.GetBytes(wordBase64); 
            
            FileStream fileStream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            document.Open(fileStream, FormatType.Doc);
            foreach (var rd in replaysDictionary)
            {
                if (string.IsNullOrEmpty(document.GetText())) continue;

                document.Replace(rd.Key, rd.Value, false, false);
                while (document.GetText().IndexOf(rd.Key) != -1)
                    document.Replace(rd.Key, rd.Value, false, false);
            }
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Doc);
            document.Close();
            stream.Position = 0; 
            return "";
        }

    }
}
