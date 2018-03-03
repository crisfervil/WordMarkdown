using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordMarkdown
{
    class Program
    {
        static void Main(string[] args)
        {

            // the path name from the command line
            var inputFileName = args.Length > 0 ? args[0] : null;

#if DEBUG
            // only for debuggin purposes
            inputFileName = String.IsNullOrEmpty(inputFileName) ? "test.docx" : inputFileName;
#endif

            var outputFileName = args.Length > 1 ? args[1] : null;
            outputFileName = string.IsNullOrEmpty(outputFileName) ? $"{inputFileName}.json" : outputFileName;

            var md = new WordMarkdown();
            var jsonData = md.GetJSon(inputFileName);
            File.WriteAllText(outputFileName, jsonData);
        }
    }
}
