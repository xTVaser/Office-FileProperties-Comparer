using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompAppsCheatingDetector {

    public class Match {

        Suspect sOne;
        Suspect sTwo;

        FileInfo fOne;
        FileInfo fTwo;

        List<string> info = new List<string>();

        public Match(Suspect one, FileInfo fOne, Suspect two, FileInfo fTwo) {

            sOne = one;
            sTwo = two;

            this.fOne = fOne;
            this.fTwo = fTwo;
        }

        public void addInfo(string msg) {

            info.Add(msg);
        }

        public Suspect getFirstSuspect() {

            return sOne;
        }

        public Suspect getSecondSuspect() {

            return sTwo;
        }
    }
}
