namespace DocxManipulator
{
    class Program
    {
        public static void Main()
        {
            var document = @"/home/microting/Documents/Test.docx";
            Modifier.SearchAndReplace(document);
//            Modifier.ConvertToPdf( @"Test-formel.docx",@"/home/microting/Documents");
//            for (int i = 1; i <= 6; i++)
//            {
//                Modifier.InsertPicture(document, $"Picture {i}");
//            }
//            Modifier.InsertPicture(document, $"Picture 1");

        }
        
    }
}