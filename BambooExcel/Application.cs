﻿using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BambooExcel
{
    public class Application
    {
        public MySqlConnection myConnection;
        private static Application _app;
        private Application()
        {

        }

        public static Application instance()
        {
            if (Application._app == null)
            {
                Application._app = new Application();
            }
            return Application._app;
        }

        public static void clear()
        {
            if (Application._app != null)
            {
                if (Application._app.myConnection!=null)
                    Application._app.myConnection.Close();
            }

        }
    }
}
