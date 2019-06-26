namespace DocxManipulator
{
    class Program
    {
        public static void Main()
        {
//            Modifier.SearchAndReplace(@"/home/microting/Documents/Test.docx");
//            Modifier.ConvertToPdf( @"FinalRapport-ALFASpecialaffald.docx",@"FinalRapport-ALFASpecialaffald.odt",
//                @"/home/microting/Documents");
            Modifier.InsertPicture(@"/home/microting/Documents/Test.docx", "Picture 1");
        }
        
    }
}