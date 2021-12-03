using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace PDFTransformer
{
    class OfficeToPDF
    {
        public string officePath { get; set; }

        /// <summary>
        /// Create a new OfficeToPDF transformer
        /// </summary>
        /// <param name="officePath">Path to 'soffice.exe' application</param>
        public OfficeToPDF(string officePath)
        {
            if (string.IsNullOrEmpty(officePath))
                throw new Exception("Office path 'soffice.exe' is required");

            this.officePath = officePath;
            StartOpenOffice();
        }

        /// <summary>
        /// Convert the inputFile to a PDF file
        /// </summary>
        /// <param name="inputFilePath">Path of the file to convert</param>
        /// <param name="outputDirPath">Path where the file will be placed after the conversion</param>
        /// <param name="timeoutMin">The operation will be cancelled after this period of time (minutes)</param>
        public void ConvertToPDF(string inputFilePath, string outputDirPath, int timeoutMin = 30)
        {
            Process p = null;
            try
            {
                var extension = Path.GetExtension(inputFilePath);

                if (ConvertExtensionToFilterType(extension) == null)
                    throw new Exception("Not valid exception: " + extension);

                string possibleOutputCommand = "";

                if (outputDirPath != null && outputDirPath.Length > 0)
                {
                    possibleOutputCommand = string.Format("--outdir {0}", outputDirPath);
                }

                var task = Task.Run(() =>
                {
                    p = new Process();
                    p.StartInfo.FileName = officePath;
                    p.StartInfo.Arguments = string.Format(@"--norestore --nofirststartwizard --headless --convert-to pdf {0} {1}", inputFilePath, possibleOutputCommand);

                    //p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.CreateNoWindow = true;

                    bool ok = p.Start();
                    if (!ok)
                    {
                        throw new Exception("Comand '--convert-to' cannot be started");
                    }
                    p.WaitForExit();
                });

                if (!task.Wait(TimeSpan.FromSeconds(5)))
                {
                    throw new TimeoutException("Timeout time exceded: " + timeoutMin + " min");
                }

                p.Dispose();
                KillOfficeProcess();
            }
            catch (Exception e)
            {
                if (p != null)
                    p.Dispose();
                KillOfficeProcess();
                KillPossibleTmpFiles(outputDirPath);
                throw e;
            }
        }

        private void StartOpenOffice()
        {
            Process[] ps = Process.GetProcessesByName("soffice.bin");
            if (ps != null)
            {
                if (ps.Length > 0)
                    return;
                else
                {

                    Process p = new Process();
                    p.StartInfo.Arguments = "--headless --nofirststartwizard";
                    p.StartInfo.FileName = officePath;

                    p.StartInfo.CreateNoWindow = false;

                    bool result = p.Start();
                    if (result == false)
                        throw new InvalidProgramException("OpenOffice failed to start.");
                }
            }
            else
            {
                throw new InvalidProgramException(string.Format(@"OpenOffice not found at: '{0}'  Is OpenOffice installed?", this.officePath));
            }
        }
        private string ConvertExtensionToFilterType(string extension)
        {
            switch (extension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                case ".css":
                case ".json":
                    return "writer_pdf_Export";
                case ".xls":
                case ".xlsb":
                case ".xlsx":
                case ".ods":
                case ".csv":
                    return "calc_pdf_Export";
                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }
        private void KillOfficeProcess()
        {
            try
            {
                Process[] ps = Process.GetProcessesByName("soffice.bin");
                if (ps != null)
                {
                    if (ps.Length > 0)
                    {
                        foreach (Process p in ps)
                        {
                            p.Kill();
                            p.WaitForExit(3000);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// En caso de forzar la salida del Office (KillOfficeProcess()) se quedarán files residuales que debemos limpiar
        /// </summary>
        /// <param name="path">Directory to check tmp files</param>
        private void KillPossibleTmpFiles(string directory)
        {
            try
            {
                if (Directory.Exists(directory))
                {
                    var allTmpFiles = Directory.EnumerateFiles(directory).Where(file => file.ToLower().EndsWith(".tmp") || file.ToLower().EndsWith(".pdf#"));
                    foreach (var file in allTmpFiles)
                    {
                        File.Delete(file);
                    }
                }
            }
            catch (Exception e) { }
        }

    }
}
