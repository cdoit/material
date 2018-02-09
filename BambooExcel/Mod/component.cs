using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    public partial class component
    {
        private String id;

        public String Id
        {
            get { return id; }
            set { id = value; }
        }

        private String projectid;

        public String Projectid
        {
            get { return projectid; }
            set { projectid = value; }
        }

        private String name;

        public String Name
        {
            get { return name; }
            set { name = value; }
        }
    }
}
