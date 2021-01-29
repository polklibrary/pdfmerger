using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace PDFMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Author: David Hietpas");
                Console.WriteLine("Email: hietpasd@uwosh.edu");
                Console.WriteLine("Updated: 1/29/2021");
                Console.WriteLine("For: UW Oshkosh Archives");
                Console.WriteLine("Description: This program will merge PDF files into a single PDF. Follow instructions below to begin.");
                Console.WriteLine("");
                Console.WriteLine("");
                Console.WriteLine("Type in the folder path to the root PDF folder OR leave blank which will scan the folder the EXE file is located, then press enter.");
                Console.Write("> ");
                string RootPath = Console.ReadLine().Replace("\\", "\\\\");
                if (RootPath == "" || RootPath == null)
                {
                    RootPath = Directory.GetCurrentDirectory();
                }

                Console.WriteLine("");
                Console.WriteLine("Type the folder name to find OR leave blank to search all, then press enter.  (Exact Search: MyFolder1) (Wildcard Search: MyFol*)");
                Console.Write("> ");
                string Target = Console.ReadLine();
                if (Target == "" || Target == null)
                {
                    Target = "*";
                }

                Console.WriteLine("");
                Console.WriteLine("Type in name of output file, then press enter. (Do not include .pdf)");
                Console.Write("> ");
                string OutputName = Console.ReadLine();
                if (OutputName == "" || OutputName == null)
                {
                    OutputName = "output";
                }
                OutputName += ".pdf";

                Console.WriteLine("");
                Console.WriteLine("Search Directory: " + RootPath);
                Console.WriteLine("Search For: " + Target);
                Console.WriteLine("Output Filename: " + OutputName);
                Console.WriteLine("Begin?  y or n, then enter.");
                Console.Write("> ");
                string Start = Console.ReadLine().ToLower();




                int DirectorySearchCount = 0;
                int PDFMergeCount = 0;
                if (Start == "y")
                {
                    using (PdfDocument FinalPDF = new PdfDocument())
                    {

                        DirectoryInfo RootInfo = new DirectoryInfo(RootPath);
                        foreach (DirectoryInfo DirInfo in RootInfo.EnumerateDirectories(Target).OrderBy(d => d.Name))
                        {
                            Console.WriteLine("Directory: " + DirInfo.Name);
                            DirectorySearchCount++;

                            foreach (FileInfo FInfo in DirInfo.EnumerateFiles().OrderBy(d => d.Name))
                            {
                                Console.WriteLine(" >> Merging " + FInfo.Name);
                                using (PdfDocument tmp = PdfReader.Open(Path.Combine(FInfo.DirectoryName, FInfo.Name), PdfDocumentOpenMode.Import))
                                {
                                    CopyPages(tmp, FinalPDF);
                                    PDFMergeCount++;
                                }
                            }
                        }

                        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                        FinalPDF.Save(Path.Combine(RootPath, OutputName));
                    }

                    Console.WriteLine("");
                    Console.WriteLine("COMPLETE---------------------------------------");
                    Console.WriteLine("Directories Searched: " + DirectorySearchCount);
                    Console.WriteLine("PDF's Merged: " + PDFMergeCount);
                    Console.WriteLine("");



                } // end start

            }
            catch (Exception e)
            {
                Console.WriteLine("Sorry, an error occurred...");
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("Press enter to close.");
            Console.Read();
        }



        static void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }
    }
}
