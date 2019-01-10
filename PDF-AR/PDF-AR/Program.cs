/*
 * PDF Archive Reader — rip the attachments out of a PDF file
 *  Copyright (C) 2018–2019 Jon Hood
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Affero General Public License as
 *  published by the Free Software Foundation, either version 3 of the
 *  License, or (at your option) any later version.

 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Affero General Public License for more details.
 *
 *  You should have received a copy of the GNU Affero General Public License
 *  along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;

namespace PDF_AR
{
    class Program
    {
        static void PrintHelp()
        {
            Console.WriteLine("Usage: PDF-AR [-h] [-i] -o=c:\\output\\directory EmbeddedAttachments.pdf");
            Console.WriteLine(" -o=[dir]       Set Output Directory");
            Console.WriteLine(" -h             This help screen");
            Console.WriteLine(" -i             Interactive Mode");
            Console.WriteLine(" [EmbeddedAttachments.pdf] PDF file with attachments");
        }

        static void Main(string[] args)
        {
            string pdfFile = "";
            string outputDir = "";
            bool interactive = false;

            foreach (string arg in args)
            {
                switch (arg.Substring(0, 2).ToLower())
                {
                    case "-o":
                        outputDir = arg.Substring(3);
                        if (!Directory.Exists(outputDir))
                            Directory.CreateDirectory(outputDir);
                        if (!Directory.Exists(outputDir))
                        {
                            Console.WriteLine("Unable to create directory: " + outputDir);
                            PrintHelp();
                            return;
                        }
                        break;
                    case "-h":
                        PrintHelp();
                        break;
                    case "-i":
                        interactive = true;
                        break;
                    default:
                        pdfFile = arg;
                        break;
                }
            }

            // read the supplied PDF
            PdfReader pdf = new PdfReader(pdfFile);
            // for every file type in the PDF in every page, build a list of tuples containing the potential attachment
            List<Tuple<string, byte[]>> files = new List<Tuple<string, byte[]>>();
            int foundItems = 0;
            for (int i = 1; i <= pdf.NumberOfPages; i++)
            {
                PdfArray arr = pdf.GetPageN(i).GetAsArray(PdfName.ANNOTS);
                if (arr == null)
                    continue;
                for (int j = 0; j < arr.Size; j++)
                {
                    //get every reference on this page and see if it's an attachment
                    PdfDictionary annot = arr.GetAsDict(j);
                    if (PdfName.FILEATTACHMENT.Equals(annot.GetAsName(PdfName.SUBTYPE)))
                    {
                        PdfDictionary fs = annot.GetAsDict(PdfName.FS);
                        PdfDictionary refs = fs.GetAsDict(PdfName.EF);
                        foreach (PdfName name in refs.Keys)
                        {
                            //found a file! Store it.
                            foundItems++;
                            string n = fs.GetAsString(name).ToString();
                            Console.WriteLine("[" + foundItems.ToString() + "] " + n);
                            files.Add(new Tuple<string, byte[]>(n, PdfReader.GetStreamBytes((PRStream)refs.GetAsStream(name))));
                        }
                    }
                }
            }
            //allow the user to select which file to save
            if (interactive)
            {
                bool ret = true;
                while (ret)
                {
                    Console.Write("Enter the number of the file you want to save (or 0 to quit): ");
                    int i = 0;
                    if (int.TryParse(Console.ReadLine(), out i))
                    {
                        if (i > 0 && i <= files.Count)
                        {
                            i--;
                            using (FileStream fs = new FileStream(Path.Combine(outputDir, files[i].Item1), FileMode.Create, FileAccess.Write))
                            {
                                fs.Write(files[i].Item2, 0, files[i].Item2.Length);
                            }
                        }
                    }
                    if (i < 0)
                        ret = false;
                }
            }
            else //or save all files in the specified directory (will overwrite)
            {
                foreach (var file in files)
                {
                    using (FileStream fs = new FileStream(Path.Combine(outputDir, file.Item1), FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(file.Item2, 0, file.Item2.Length);
                    }
                }
            }
        }
    }
}
