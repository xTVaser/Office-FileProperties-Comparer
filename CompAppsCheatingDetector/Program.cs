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
                    suspectIterator(suspects);
                }

                Console.WriteLine("");
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        static void suspectIterator(List<Suspect> suspects) {

            for (int i = 0; i < suspects.Count-1; i++) { //For each suspect except the last one
                
                foreach(FileInfo f in suspects[i].getFiles()) { //That suspects file lists.

                    fileIterator(f, i + 1, suspects, suspects[i]); //Combine each parent file to every other parent file.
                }
            }

        }

        static void fileIterator(FileInfo f1, int index, List<Suspect> suspects, Suspect s) {

            for(int i = index;  i < suspects.Count; i++) {

                foreach(FileInfo f2 in suspects[i].getFiles()) {

                    if (checkExtension(f1) == FILETYPE_OTHER || checkExtension(f2) == FILETYPE_OTHER) //If either are wrong 
                        matchNormalFiles(f1, f2, s, suspects[i]);
                    else if (checkExtension(f1) != FILETYPE_OTHER && checkExtension(f1) == checkExtension(f2)) //If they are both office files
                        matchOfficeFiles(f1, f2, s, suspects[i], checkExtension(f1));
                }
            }
        }

        readonly static int FILETYPE_WORD = 1;
        readonly static int FILETYPE_EXCEL = 2;
        readonly static int FILETYPE_PPT = 3;
        readonly static int FILETYPE_OTHER = 4;

        static int checkExtension(FileInfo file) {
            
            if (file.Extension.Equals(".docx") || file.Extension.Equals(".doc")) //Its a word document
                return FILETYPE_WORD;

            else if (file.Extension.Equals(".xls") || file.Extension.Equals(".xlm") || file.Extension.Equals(".xlsx")) //Its an excel document
                return FILETYPE_EXCEL;

            else if (file.Extension.Equals(".ppt") || file.Extension.Equals(".pptx")) //Its a powerpoint document
                return FILETYPE_PPT;

            return FILETYPE_OTHER;
        }

        static void matchNormalFiles(FileInfo f1, FileInfo f2, Suspect s1, Suspect s2) {

            List<string> matchInfo = new List<string>();
            matchInfo.Add("Normal File Properties:");

            //File Info Properties
            if (f1.CreationTime != null && f2.CreationTime != null && f1.CreationTime.Equals(f2.CreationTime))
                matchInfo.Add("Same Creation Time - File 1:" + f1.CreationTime + " File 2:" + f2.CreationTime);
            if (f1.LastAccessTime != null && f2.LastAccessTime != null && f1.LastAccessTime.Equals(f2.LastAccessTime))
                matchInfo.Add("Same Last Access Time - File 1:" + f1.LastAccessTime + " File 2:" + f2.LastAccessTime);
            if (f1.LastWriteTime != null && f2.LastWriteTime != null && f1.LastWriteTime.Equals(f2.LastWriteTime))
                matchInfo.Add("Same Last Write Time - File 1: " + f1.LastWriteTime + " File 2:" + f2.LastWriteTime);
            if (f1.Length == f2.Length)
                matchInfo.Add("Same File Length - File 1: " + f1.Length + " File 2:" + f2.Length);
            if (f1.Name != null && f2.Name != null && f1.Name.Equals(f2.Name))
                matchInfo.Add("Same Name - File 1: " + f1.Name + " File 2:" + f2.Name);

            if (matchInfo.Count > 1) {

                Match newMatch = new Match(s1, f1, s2, f2);

                foreach(string s in matchInfo)
                    newMatch.addInfo(s);

                s1.addMatch(newMatch);
                s2.addMatch(newMatch);
            }
        }
        
        static void matchOfficeFiles(FileInfo f1, FileInfo f2, Suspect s1, Suspect s2, int type) {

            List<string> matchInfo = new List<string>();
            matchInfo.Add("Office Document Properties:");

            OpenXmlPackage file1;
            OpenXmlPackage file2;

            if(type == FILETYPE_WORD) {
                file1 = WordprocessingDocument.Open(f1.FullName.Replace("\\", "/"), false);
                file2 = WordprocessingDocument.Open(f2.FullName.Replace("\\", "/"), false);
            }
            else if (type == FILETYPE_EXCEL) {
                file1 = SpreadsheetDocument.Open(f1.FullName.Replace("\\", "/"), false);
                file2 = SpreadsheetDocument.Open(f2.FullName.Replace("\\", "/"), false);
            }
            else {
                file1 = PresentationDocument.Open(f1.FullName.Replace("\\", "/"), false);
                file2 = PresentationDocument.Open(f2.FullName.Replace("\\", "/"), false);
            }

            var props1 = file1.PackageProperties;
            var props2 = file2.PackageProperties;

            //Office File Properties
            if (props1.Category != null && props2.Category != null && props1.Category.Equals(props2.Category))
                matchInfo.Add("Same Category - File 1:" + props1.Category + " File 2:" + props2.Category);
            if (props1.ContentStatus != null && props2.ContentStatus != null && props1.ContentStatus.Equals(props2.ContentStatus))
                matchInfo.Add("Same Content Status - File 1:" + props1.ContentStatus + " File 2:" + props2.ContentStatus);
            if (props1.ContentType != null && props2.ContentType != null && props1.ContentType.Equals(props2.ContentType))
                matchInfo.Add("Same Content Type - File 1:" + props1.ContentType + " File 2:" + props2.ContentType);
            if (props1.Created != null && props2.Created != null && props1.Created.Equals(props2.Created))
                matchInfo.Add("Same Date Created - File 1:" + props1.Created + " File 2:" + props2.Created);
            if (props1.Creator != null && props2.Creator != null && props1.Creator.Equals(props2.Creator))
                matchInfo.Add("Same Creator - File 1:" + props1.Creator + " File 2:" + props2.Creator);
            if (props1.Description != null && props2.Description != null && props1.Description.Equals(props2.Description))
                matchInfo.Add("Same Description - File 1:" + props1.Description + " File 2:" + props2.Description);
            if (props1.Identifier != null && props2.Identifier != null && props1.Identifier.Equals(props2.Identifier))
                matchInfo.Add("Same Identifier - File 1:" + props1.Identifier + " File 2:" + props2.Identifier);
            if (props1.Keywords != null && props2.Keywords != null && props1.Keywords.Equals(props2.Keywords))
                matchInfo.Add("Same Keywords - File 1:" + props1.Keywords + " File 2:" + props2.Keywords);
            if (props1.LastModifiedBy != null && props2.LastModifiedBy != null && props1.LastModifiedBy.Equals(props2.LastModifiedBy))
                matchInfo.Add("Same Last Modified By - File 1:" + props1.LastModifiedBy + " File 2:" + props2.LastModifiedBy);
            if (props1.LastPrinted != null && props2.LastPrinted != null && props1.LastPrinted.Equals(props2.LastModifiedBy))
                matchInfo.Add("Same Last Printed - File 1:" + props1.LastPrinted + " File 2:" + props2.LastPrinted);
            if (props1.Modified != null && props2.Modified != null && props1.Modified.Equals(props2.Modified))
                matchInfo.Add("Same Modified Date - File 1:" + props1.Modified + " File 2:" + props2.Modified);
            if (props1.Revision != null && props2.Revision != null && props1.Revision.Equals(props2.Revision))
                matchInfo.Add("Same Revision - File 1:" + props1.Revision + " File 2:" + props2.Revision);
            if (props1.Subject != null && props2.Subject != null && props1.Subject.Equals(props2.Subject))
                matchInfo.Add("Same Subject - File 1:" + props1.Subject + " File 2:" + props2.Subject);
            if (props1.Title != null && props2.Title != null && props1.Title.Equals(props2.Title))
                matchInfo.Add("Same Title - File 1:" + props1.Title + " File 2:" + props2.Title);
            if (props1.Version != null && props2.Version != null && props1.Version.Equals(props2.Version))
                matchInfo.Add("Same Version - File 1:" + props1.Version + " File 2:" + props2.Version);

            //File Info Properties
            if (f1.CreationTime != null && f2.CreationTime != null && f1.CreationTime.Equals(f2.CreationTime))
                matchInfo.Add("Same Creation Time - File 1:" + f1.CreationTime + " File 2:" + f2.CreationTime);
            if (f1.LastAccessTime != null && f2.LastAccessTime != null && f1.LastAccessTime.Equals(f2.LastAccessTime))
                matchInfo.Add("Same Last Access Time - File 1:" + f1.LastAccessTime + " File 2:" + f2.LastAccessTime);
            if (f1.LastWriteTime != null && f2.LastWriteTime != null && f1.LastWriteTime.Equals(f2.LastWriteTime))
                matchInfo.Add("Same Last Write Time - File 1: " + f1.LastWriteTime + " File 2:" + f2.LastWriteTime);
            if (f1.Length == f2.Length)
                matchInfo.Add("Same File Length - File 1: " + f1.Length + " File 2:" + f2.Length);
            if (f1.Name != null && f2.Name != null && f1.Name.Equals(f2.Name))
                matchInfo.Add("Same Name - File 1: " + f1.Name + " File 2:" + f2.Name);

            if (matchInfo.Count > 1) {

                Match newMatch = new Match(s1, f1, s2, f2);

                foreach (string s in matchInfo)
                    newMatch.addInfo(s);

                s1.addMatch(newMatch);
                s2.addMatch(newMatch);
            }
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

            MessageBox.Show("Found " + suspects.Count.ToString()+" Suspects", "Message");

            return suspects;
        }
    }
}
