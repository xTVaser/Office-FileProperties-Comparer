﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CompAppsCheatingDetector {
    public class Suspect {

        string name;
        List<FileInfo> fileList = new List<FileInfo>();
        List<Match> matchList = new List<Match>();

        public Suspect(string name) {
            this.name = name;
        }

        public void addFile(FileInfo file) {

            fileList.Add(file);
        }

        public void addMatch(Match match) {

            matchList.Add(match);
        }

        public string getName() {
            return name;
        }

        public List<Match> getMatches() {
            return matchList;
        }

        public List<FileInfo> getFiles() {
            return fileList;
        }
        
        public override string ToString() {

            return name + " Files: " + fileList.Count.ToString() + " Matches: "+matchList.Count.ToString();
        }
    }
}
