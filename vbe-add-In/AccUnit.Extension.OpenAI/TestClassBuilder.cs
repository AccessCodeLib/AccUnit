using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenAI;

namespace AccessCodeLib.AccUnit.Extension.OpenAI
{
    public class TestClassBuilder
    {
        public string BuildVbaCode()
        {
            var client = new OpenAIClient(Environment.GetEnvironmentVariable("OPENAI_API_KEY"));
            

            return "";
        }


    }
}
