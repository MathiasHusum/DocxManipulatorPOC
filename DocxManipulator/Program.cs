using System;
namespace DocxManipulator
{
    class Program
    {
        public static void Main(string[] args)
        {
            Modifier.SearchAndReplace(@"/home/microting/Documents/FinalRapport-ALFASpecialaffald.docx");
            Modifier.ConvertToPdf( @"FinalRapport-ALFASpecialaffald.docx",@"FinalRapport-ALFASpecialaffald.odt",
                @"/home/microting/Documents");
        }
        
    }
}