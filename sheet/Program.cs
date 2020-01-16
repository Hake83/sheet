using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace sheet
{
    class Program
    {
        static void Main()
        {
            var fi = new FileInfo(@"P:\Engineering\Engineering Schedule.xlsm");
            // Directory object for searching through files in the current directory
            string dir = Directory.GetCurrentDirectory();

            // If the file isn't found this will be an array of zero size
            string[] files = Directory.GetFiles(dir, "N*.pdf", SearchOption.TopDirectoryOnly);

            if (files.Count() > 0)
            {
                //Run Data function to get data needing to go into schedule file
                string[] info = Data(files[0]);

                //Data returns the hours in a string with special characters that needs to be converted to a more useable int
                string[] hour = info[4].Split(new char[] { ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
                int hs = 0;
                if (info[4] != " ")
                {
                    hs = Convert.ToInt32(info[4]);
                    hs = hs / 100000;
                }

                //Get the parent directory name to fill the job number column
                string parentDir = Directory.GetCurrentDirectory().ToString() == null ? null : Directory.GetParent(dir).ToString();

                if (parentDir != null)
                {
                    parentDir = System.IO.Path.GetFileName(parentDir);
                }

                //Call function to find last line
                int b = FirstLine(fi);

                //Opens engineering schedule and fills spreadsheet cells if possible
                try
                 {
                     var p = new ExcelPackage(fi);
                     ExcelWorksheet ws = p.Workbook.Worksheets[1];
                    bool test = true;
                    for (int n = 4; n<=b-1; n++)
                    {
                        if (ws.Cells[n, 5].Value.ToString() == info[1].ToString())
                        {
                            Console.WriteLine("This job is already on the engineering schedule");
                            test = false;
                            break;
                        }
                    }
                    string[] jobNumber = parentDir.Split(new char[] { ' ', '-' });

                    if (test==true)
                    {
                        ws.Cells[b, 6].Value = info[0];
                        ws.Cells[b, 5].Value = info[1];
                        ws.Cells[b, 10].Style.Numberformat.Format = "yyyy-mm-dd;@";
                        ws.Cells[b, 10].Value = info[2];
                        ws.Cells[b, 7].Value = "ECI " + info[3];
                        ws.Cells[b, 8].Value = hs;
                        ws.Cells[b, 11].Value = "Not Assigned";
                        ws.Cells[b, 14].Value = "No";
                        ws.Cells[b, 15].Value = "Yes";
                        try
                        {
                            //ws.Cells[b, 4].Value = jobNumber[0] + "-" + jobNumber[1];  Commented out because new job number format (A00020 doesn't include the 000/001 and was breaking)
                            ws.Cells[b, 4].Value = jobNumber[0];
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Something was wrong with directory formatting to get the work order, fill this manually");
                        }

                        ws.Cells[b, 9].Formula = "IF(J" + b + "=\"\",\"\",J"+b+"-21)";
                        ws.Calculate();
                        p.Save();
                        Console.WriteLine("Writing to Engineering Schedule...");
                    }
                 }

                 catch (IOException)
                 {
                     Console.WriteLine("It appears the engineering schedule file is open please close it and retry the program");
                 }

                 catch (Exception error)
                 {
                     Console.WriteLine("something went wrong");
                     Console.WriteLine(error);
                     Console.ReadKey();
                 }

                //Copy RFQ files
                try
                {
                    //string[] sourceDir = Directory.GetDirectories(@"P:\Office\Quotes\2018\Honeywell\Tooling Logistics\Design Quotes\", "*" + info[0] + "*");
                    string[] sourceDir = Directory.GetDirectories(@"P:\Office\Quotes\2019\Honeywell\Tooling Logistics\", "*" + info[5] + "*", SearchOption.AllDirectories);
                    //string[] fileList = Directory.GetFiles(dirQuotes[0]);

                    FileInfo file = new FileInfo(Directory.GetParent(dir) + @"\Drawings\Customer Supplied\Customer Supplied\");
                    file.Directory.Create();
                    string destination = Directory.GetParent(dir) + @"\Drawings\Customer Supplied\Customer Supplied\";
                    string[] check = Directory.GetFiles(destination);
                    if (check.Count() == 0)
                    {
                        DirectoryCopy(sourceDir[0], destination, true);
                        Console.WriteLine("Copying Customer Drawing Files...");
                    }
                    else
                    {
                        Console.WriteLine("There are already files in the Customer Drawing folder");
                    }
                }
                catch (Exception error)
                {
                    Console.WriteLine("Couldn't find the quote directory you'll have to move stuff over manually");
                    Console.WriteLine(error);
                    Console.ReadKey();
                }
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("No PO file found check filename should start with 'N'");
                Console.WriteLine("Are you in the PO folder?");
                Console.ReadKey();
            }
        }

        // Function to return first usable row in the worksheet based on PO number column
        static int FirstLine(FileInfo a)
        {
            using (var p = new ExcelPackage(a))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1];

                int n = 0;
                int i = 4;
                while (n < 1)
                {
                    if (ws.Cells[i, 4].Value != null)
                    {
                        i++;
                    }
                    else n++;
                }
                return i;
            }
        }

        //Function to get information from the Purchase Order
        static string[] Data(string b)
        {
            PdfReader reader = new PdfReader(b);
            string pn = PdfTextExtractor.GetTextFromPage(reader, 1);
            string[] po = pn.Split(new char[] { ' ', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            reader.Close();

            string partNumber = " ";
            string searchNumber = " ";
            //Find the partNumber
            int n = 0;
            while (n <= po.Count())
            {
                if (po[n] == "Design" || po[n] == "design")
                {
                    partNumber = po[n + 1] + "-" + po[n + 3];
                    searchNumber = po[n + 1] + "*" + po[n + 3];
                    break;
                }
                if (n >= po.Count() - 1)
                {
                    partNumber = "Fail";
                    searchNumber = "Fail";
                    break;
                }
                else
                {
                    n++;
                }
            }

            // Find the Purchase Order number
            string purchaseNumber = " ";
            Match m = Regex.Match(pn, @"N\d+");
            purchaseNumber = m.Value;

            string dueDate = Regex.Matches(pn, @"\d{2}\/\d{2}\/\d{4}")[1].Value;
            DateTime date = DateTime.Parse(dueDate);
            dueDate = date.ToString("MM/dd/yyyy");

            // Find eci yes or no
            n = 0;
            string eci = " ";
            while (n <= po.Count())
            {
                if (po[n] == "ECI:" || po[n] == "eci:" || po[n] == "ECI" || po[n] == "eci")
                {
                    eci = po[n + 1];
                    break;
                }
                if (n >= po.Count() - 1)
                {
                    break;
                }
                else
                {
                    n++;
                }
            }

            // Find price
            Match mhour = Regex.Match(pn, @"\d?\d?\,?\d+\.\d{3}");
            string hours = mhour.Value;

            if (mhour.Success == true)
            {
                string[] h = hours.Split(new char[] { ',', '.' }, StringSplitOptions.RemoveEmptyEntries);

                int i = 0;
                hours = "";
                while (i < h.Count())
                {
                    hours = hours + h[i];
                    i++;
                }
            }
            else
            {
                hours = " ";
            }
            string[] stuff = { partNumber, purchaseNumber, dueDate, eci, hours, searchNumber };
            return stuff;
        }

        //Function to fill excel sheet Engineering Schedule (Not used at this time)
        static void ScheduleFill(string[] data, int line, int hour, FileInfo wb)
        {
            try
            {
                var p = new ExcelPackage(wb);

                ExcelWorksheet ws = p.Workbook.Worksheets[1];

                ws.Cells[line, 6].Value = data[0];
                Console.WriteLine(ws.Cells[line, 6].Value);
                Console.WriteLine(line);
                ws.Cells[line, 5].Value = data[1];
                ws.Cells[line, 10].Value = data[2];
                ws.Cells[line, 7].Value = "ECI: " + data[3];
                ws.Cells[line, 8].Value = hour;
                ws.Cells[line, 11].Value = "Not Assigned";
                ws.Cells[line, 14].Value = "No";
                ws.Cells[line, 15].Value = "Yes";
                p.Save();

            }
            catch (IOException e)
            {
                Console.WriteLine("Engineering Schedule appears to be open, please close it and retry this program");
                Console.WriteLine(e);
            }
            catch (Exception error)
            {
                Console.WriteLine("Something went wrong");
                Console.WriteLine(error);
            }
        }

        // Function to copy directory and contents, stolen from msdn.microsoft.com
        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = System.IO.Path.Combine(destDirName, file.Name);

                // Copy the file.
                file.CopyTo(temppath, false);
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = System.IO.Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
    }
}