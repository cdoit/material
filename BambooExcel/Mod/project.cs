using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    public partial class project
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

        private String code;

        public String Code
        {
            get { return code; }
            set { code = value; }
        }

        private String address;

        public String Address
        {
            get { return address; }
            set { address = value; }
        }

        private String area;

        public String Area
        {
            get { return area; }
            set { area = value; }
        }

    }
}
