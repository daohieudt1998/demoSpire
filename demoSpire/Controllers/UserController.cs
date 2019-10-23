using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using demoSpire.Helper;
using DigitalBilling;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Pdf;

namespace demoSpire.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserController : ControllerBase
    {
        private Digital _digital;
        public UserController(Digital digital)
        {
            _digital = digital;
        }
        [HttpGet]
        [Route("word")]
        //Get: api/User/word
        public IActionResult addWord()
        {
            Document doc = new Document();
            doc.LoadFromFile("D:/Demo/demoSpire/demoSpire/Content/test.docx");
            doc.Replace("Document", "daohieu", true, true);
            doc.SaveToFile("daohieu.docx", Spire.Doc.FileFormat.Docx2013);
            return Ok();
        }
        
        [HttpGet]
        [Route("createBilling")]
        //Get: api/User/createBilling
        public IActionResult createBilling()
        {
            string a = _digital.create1();
            return Ok(a);
        }

        [HttpGet]
        [Route("makeSearchID")]
        //Get: api/User/makeSearchID
        public string makeSearchID()
        {
            long i = 1;
            foreach (byte b in Guid.NewGuid().ToByteArray())
            {
                i *= ((int)b + 1);
            }
            string newId = string.Format("{0:x}", i - DateTime.Now.Ticks);
            return newId;
        }

        [HttpGet]
        [Route("convertNumber")]
        // Post: api/User/convertNumber
        public IActionResult convertNumber()
        {
            var n = 6919191919;
            ConvertToCurrency convert = new ConvertToCurrency();
            string output = convert.NumberToTextVN(n);
            return Ok(output);
        }

        [HttpGet]
        [Route("test")]
        // Post: api/User/test
        public IActionResult test()
        {
            string startupPath = Environment.CurrentDirectory;
            return Ok(startupPath);
        }
    }
}