using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;

namespace WordToPDF
{
    class WordToPDF
    {
        // Fields
        private FileFormat m_fileFormat = FileFormat.PDF;

        private string m_pathToFile;
        private string m_fileName;
        private string[] m_files = { };

        // Properties
        public string FilePath
        {
            get { return m_pathToFile; }
            set 
            {
                if (Directory.Exists(value))
                    m_pathToFile = value;
                else
                    throw new DirectoryNotFoundException("The directory specified has not been found!");
            }
        }

        public string FileName
        {
            get { return m_fileName; }
            set 
            {
                    m_fileName = value;
            }
        }


        // Constructors
        public WordToPDF()
        {
        }

        public WordToPDF(string pathToFile, string fileName)
        {

            m_fileName = fileName;
            m_pathToFile = pathToFile;
        }

        /// <summary>
        /// Runs the whole script to convert the files to PDF.
        /// </summary>
        public void RunConversions()
        {
            runChecks();

            if (m_files.Length != 0)
            {
                foreach (string file in m_files)
                {
                    if (Path.GetExtension(file) != ".pdf")
                        Convert(m_pathToFile, file);
                }
            }
            if (!string.IsNullOrEmpty(m_fileName))
            {
                Convert(m_pathToFile, m_fileName);
            }
        }

        /// <summary>
        /// Checks if the user has inputed a filename, a path or a path with the file name.
        /// </summary>
        private void runChecks()
        {
            if (File.Exists(m_pathToFile))
            {
                Console.WriteLine($"Converting {Path.GetFileName(m_pathToFile)} to {Path.GetFileNameWithoutExtension(m_pathToFile)}.pdf!");
                m_fileName = Path.GetFileNameWithoutExtension(m_pathToFile);
                m_pathToFile = Path.GetFullPath(m_pathToFile);
            }
            else if (Directory.Exists(m_pathToFile))
            {
                if (string.IsNullOrEmpty(m_fileName))
                { 
                    m_files = Directory.GetFiles(m_pathToFile);
                    Console.WriteLine("You have given a path as an argument without specifying the file. All files will now be converted to PDF.");
                }

            }

        }


        /// <summary>
        /// Convert files from a type to PDF
        /// </summary>
        private void Convert(string path, string file)
        {
            try
            {
                string fullPath = Path.Combine(path, file);
                Console.WriteLine($"Loading file: {file}");
                Document doc = new Document();
                doc.LoadFromFile(@$"{fullPath}");
                string fileName = Path.GetFileNameWithoutExtension(file);
                doc.SaveToFile($"{Path.Combine(path, fileName)}.pdf", m_fileFormat);
                Console.WriteLine("Document created!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void HelpMessage()
        {
            Console.WriteLine("Welcome to .Net Word to PDF converter. Here is how to use it:");
            Console.WriteLine("");
            Console.WriteLine(@"Example: Convert -file 'C:\path\to\file.pdf'");
            Console.WriteLine(@"Example: Convert -path 'C:\path\to\' -file file.pdf");
            Console.WriteLine(@"Example: Convert -path 'C:\path\to\'");
            Console.WriteLine("");
            Console.WriteLine("\t[ + ] -file => Path to the file, including the file.");
            Console.WriteLine("");
            Console.WriteLine("\t[ + ] -path => Path to the file.");
            Console.WriteLine("\t\t|--> Excluding the file name (-file file.pdf) will convert to PDF all of the files in the directory.");
            Console.WriteLine("");
            Console.WriteLine("\t[ + ] -h or --help => Display this help message");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine("Press ENTER to exit.");
            Console.ReadKey();
        }
    }
}
