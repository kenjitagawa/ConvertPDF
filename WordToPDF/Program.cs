using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;




namespace WordToPDF
{
    class Program
    {
        static void Main(string[] args)
        {


            // Parameters
            WordToPDF converter = new WordToPDF();

            string fileName;
            string directory;

            if (args.Length == 0)
            {
                converter.HelpMessage();
                return;
            }
            else if (args[0] == "-h" || args[0] == "-help")
            {
                converter.HelpMessage();
                return;
            }

            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(args[i]);
                if (args[i] == "-file")
                {
                    fileName = args[i + 1];
                    converter.FileName = fileName;
                }
                else if (args[i] == "-path")
                {
                    directory = args[i + 1];
                    converter.FilePath = directory;
                }
                
            }

            

            converter.RunConversions();



        }
    }
}
