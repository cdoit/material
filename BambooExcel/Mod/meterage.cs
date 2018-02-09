using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    public partial class meterage
    {
        private String id;

        public String Id
        {
            get { return id; }
            set { id = value; }
        }

        private String parentid;

        public String Parentid
        {
            get { return parentid; }
            set { parentid = value; }
        }

        private String componentid;

        public String Componentid
        {
            get { return componentid; }
            set { componentid = value; }
        }

        private String materielcode;

        public String Materielcode
        {
            get { return materielcode; }
            set { materielcode = value; }
        }

        private String materielname;

        public String Materielname
        {
            get { return materielname; }
            set { materielname = value; }
        }

        private String specifications;

        public String Specifications
        {
            get { return specifications; }
            set { specifications = value; }
        }

        private String count;

        public String Count
        {
            get { return count; }
            set { count = value; }
        }

        private String unit;

        public String Unit
        {
            get { return unit; }
            set { unit = value; }
        }

        private String rule;

        public String Rule
        {
            get { return rule; }
            set { rule = value; }
        }

        private String lossrate;

        public String Lossrate
        {
            get { return lossrate; }
            set { lossrate = value; }
        }


        private String memo;

        public String Memo
        {
            get { return memo; }
            set { memo = value; }
        }

    }
}
