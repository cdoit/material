
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
namespace IronPython_TestDll
{
    public class TestDll
    {
        public  string Add(string json)
        {        
            CDO.BaiJuyi.Tym.TymCommand command = new CDO.BaiJuyi.Tym.TymCommand();
            var ms = command.Excute(json);
            return ms;
        }
    }


    public class TestDll1
    {
        private int aaa = 11;
        public int AAA
        {
            get { return aaa; }
            set { aaa = value; }
        }
        public void ShowAAA()
        {
            
        }


    }
}