using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatisticalTools.po
{
    class Color
    {
        private String cid;
        private String mcolor;
        public String Cid {
            set;
            get;
        }

        public String Mcolor {
            get { return mcolor; }
            set { mcolor = value; }
        }

    }
}
