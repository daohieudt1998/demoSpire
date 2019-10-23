using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace demoSpire.Helper
{
    public class MakeCodeSearchBilling
    {
        public string codeSearch()
        {
            long i = 1;
            foreach (byte b in Guid.NewGuid().ToByteArray())
            {
                i *= ((int)b + 1);
            }
            string newId = string.Format("{0:x}", i - DateTime.Now.Ticks);
            return newId;
        }
    }
}
