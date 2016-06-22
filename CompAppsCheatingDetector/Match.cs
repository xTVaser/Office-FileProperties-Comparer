using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompAppsCheatingDetector {

    public class Match {

        public Suspect sOne { get; set; }
        public Suspect sTwo { get; set; }
    
        public FileInfo fOne { get; set; }
        public FileInfo fTwo { get; set; }

        public List<string> info { get; set; } = new List<string>();

        public Match(Suspect one, FileInfo fOne, Suspect two, FileInfo fTwo) {

            sOne = one;
            sTwo = two;

            this.fOne = fOne;
            this.fTwo = fTwo;
        }

        public void addInfo(string msg) {

            info.Add(msg);
        }
    }
}
