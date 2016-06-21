using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml.Linq;

//Must install Open XML 2.0 SDK

namespace CompAppsCheatingDetector {

    class Program {

        [STAThread]
        static void Main(string[] args) {

            //Find all files in the folder, and create a list of suspects with their respective files.

            try {

                FolderBrowserDialog fbd = new FolderBrowserDialog();
                DialogResult folder = fbd.ShowDialog();
                List<Suspect> suspects;

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath)) {
                    Console.WriteLine(fbd.SelectedPath);
                    DirectoryInfo dir = new DirectoryInfo(fbd.SelectedPath);
                    FileInfo[] files = dir.GetFiles("*");
                    suspects = parseSuspects(new List<FileInfo>(files));
                }

                WordprocessingDocument document = WordprocessingDocument.Open("C:/Users/Dtylan/Desktop/test2.docx", false); //false = not editable
                
                //SpreadsheetDocument.Open("FILENAME", false);
                //PresentationDocument.Open("FILENAME", false);

                FileInfo oFileInfo = new FileInfo("Yes"); //For other windows file stuff
               

                var props = document.PackageProperties;

                Console.WriteLine("Creator: " + props.Creator);
                Console.WriteLine("Time: " + props.Modified);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        static void suspectIterator(Suspect s) {



        }
        static List<Suspect> parseSuspects(List<FileInfo> files) {
           
            List<Suspect> suspects = new List<Suspect>();

            while(files.Count != 0) { //Loop through all the files.

                string name = files[0].Name.Split('_')[0];

                Suspect currentSuspect = null;

                foreach(Suspect s in suspects) { //Check if this is another file for the same person

                    if (s.getName().Equals(name))
                        currentSuspect = s;
                }

                if (currentSuspect == null) {

                    currentSuspect = new Suspect(name);
                    suspects.Add(currentSuspect);
                }

                currentSuspect.addFile(files[0]);
                files.RemoveAt(0);

            }

            MessageBox.Show("Files found: " + files.Count.ToString(), "Message");

            return suspects;/
        }

        static void compareOfficeDocuments(OpenXmlPackage document) {
            
            Console.WriteLine(file.Extension);
            if (file.Extension.Equals(".docx") || file.Extension.Equals(".doc") { //Its a word document

            }

            else if (file.Extension.Equals(".xls") || file.Extension.Equals(".xlm") || file.Extension.Equals(".xlsx")) { //Its an excel document

            }

            else if (file.Extension.Equals(".ppt") || file.Extension.Equals(".pptx")) { //Its a powerpoint document

            }

            else {


            }
        }
    }
}
