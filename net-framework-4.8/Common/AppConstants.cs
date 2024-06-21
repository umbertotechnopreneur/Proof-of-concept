using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocLocationFinder.Common
{
    public static class AppConstants
    {
        public static class Action
        {
            public const string FindWordFileCoordinates = "wordxy";
            public const string ReplaceWordFileText = "wordremove";
            public const string ConvertWordToPdf = "wordtopdf";
            public const string MergePDFDocuments = "pdfmerge";

        }
    }
}
