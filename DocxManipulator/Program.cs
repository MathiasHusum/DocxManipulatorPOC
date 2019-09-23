namespace DocxManipulator
{
    class Program
    {
        public static void Main()
        {
            var document = @"/home/microting/Documents/Test.docx";
            Modifier.SearchAndReplace(document);
//            Modifier.ConvertToPdf( @"Test-formel.docx",@"/home/microting/Documents");
//            Modifier.Insert(document, "");
            
//            Modifier.InsertPicture(document, $"Picture 1");

        }
        
    }
}