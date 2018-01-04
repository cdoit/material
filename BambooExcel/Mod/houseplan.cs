using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel.Mod
{
    public partial class houseplan
    {
        private String id;

        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        private String houseid;

        public String Houseid
        {
            get { return houseid; }
            set { houseid = value; }
        }
        private String planname;

        public String Planname
        {
            get { return planname; }
            set { planname = value; }
        }
        private String room;

        public String Room
        {
            get { return room; }
            set { room = value; }
        }
        private String part;

        public String Part
        {
            get { return part; }
            set { part = value; }
        }
        private String materialpackge;

        public String Materialpackge
        {
            get { return materialpackge; }
            set { materialpackge = value; }
        }
    }
}
