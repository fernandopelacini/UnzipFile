using Aspose.Zip.Rar;
using Microsoft.Extensions.Configuration;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;

namespace UnzipFile
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to directory of files to compress and decompress.
            var builder = new ConfigurationBuilder().AddJsonFile($"Config.json", true, true);
            var config = builder.Build();
            var unzipPath = config["UnzipPath"];
            var zipPath = config["ZipPath"];
            var zipFileName = config["ZipFileName"];
            string fullPathName = zipPath + zipFileName;
            //string dirpath = @"c:\users\public\reports";

            //DirectoryInfo di = new DirectoryInfo(zipPath);
            FileInfo fi = new FileInfo(fullPathName);
            //Decompress(fi, unzipPath);
            //DecompressRar(fi, unzipPath);
            DecompressWinrar(fi, unzipPath);
            //// Compress the directory's files.
            //foreach (FileInfo fi in di.GetFiles())
            //{
            //    Compress(fi);
            //}

            // Decompress all *.gz files in the directory.
            //foreach (FileInfo fi in di.GetFiles(fullPathName))
            //{

            //    Decompress(fi);
            //}
        }

        public static void Compress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Prevent compressing hidden and 
                // already compressed files.
                if ((File.GetAttributes(fi.FullName)
                    & FileAttributes.Hidden)
                    != FileAttributes.Hidden & fi.Extension != ".gz")
                {
                    // Create the compressed file.
                    using (FileStream outFile =
                                File.Create(fi.FullName + ".gz"))
                    {
                        using (GZipStream Compress =
                            new GZipStream(outFile,
                            CompressionMode.Compress))
                        {
                            // Copy the source file into 
                            // the compression stream.
                            inFile.CopyTo(Compress);

                            Console.WriteLine("Compressed {0} from {1} to {2} bytes.",
                                fi.Name, fi.Length.ToString(), outFile.Length.ToString());
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Extension not valid, file extension is {0}", fi.Extension);
                }
            }
        }

        public static void Decompress(FileInfo fi, string outputPath)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Get original file extension, for example
                // "doc" from report.doc.gz.
                string curFile = fi.FullName;
                string origName = curFile.Remove(curFile.Length -
                        fi.Extension.Length);

                //Create the decompressed file.
                using (FileStream outFile = File.Create(origName))
                {

                    ZipFile.ExtractToDirectory(fi.FullName, outputPath, true);

                    Console.WriteLine("Decompressed: {0}", fi.Name);

                    //using (GZipStream Decompress = new GZipStream(inFile,
                    //    CompressionMode.Decompress))
                    //{
                    //    // Copy the decompression stream 
                    //    // into the output file.
                    //    Decompress.CopyTo(outFile);

                    //    Console.WriteLine("Decompressed: {0}", fi.Name);
                    //}
                }
            }
        }

        /// <summary>
        /// Decompress using the Aspose.zip nuget, works only with .rar extension
        /// </summary>
        /// <param name="fi"></param>
        /// <param name="outputPath"></param>
        public static void DecompressRar(FileInfo fi, string outputPath)
        {
            // Load input RAR file.
            RarArchive archive = new RarArchive(fi.FullName);

            // Unrar or extract all files from the archive
            archive.ExtractToDirectory(outputPath);
        }

        /// <summary>
        /// Decompress calling winrar.exe directly, has to be installed on terminal
        /// </summary>
        /// <param name="fi"></param>
        /// <param name="outputPath"></param>
        public static void DecompressWinrar(FileInfo fi, string outputPath)
        {
                        
            ProcessStartInfo startInfo = new ProcessStartInfo("C:\\Program Files\\WinRAR\\WinRAR.exe");
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = string.Format("x -ibck -o+ \"{0}\" \"{1}\"",
                                  fi.FullName, outputPath);

            Console.WriteLine("Startinfo arguments: " + startInfo.Arguments.ToString());
            try
            {
                // Start the process with the info we specified.

                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
