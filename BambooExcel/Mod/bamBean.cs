using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    class bamBean
    {
        private String id;

        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        private String name;

        public String Name
        {
            get { return name; }
            set { name = value; }
        }

        private String parentid;

        public String Parentid
        {
            get { return parentid; }
            set { parentid = value; }
        }
    }
}
