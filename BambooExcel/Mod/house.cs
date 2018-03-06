using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    public partial class  house
    {
        private String id;

        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        private String housename;

        public String Housename
        {
            get { return housename; }
            set { housename = value; }
        }

        private String designdata;

        public String Designdata
        {
            get { return designdata; }
            set { designdata = value; }
        }
    }
}
