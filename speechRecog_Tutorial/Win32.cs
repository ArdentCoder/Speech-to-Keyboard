using System;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Interceptor;
namespace speechRecog_Tutorial
{
   

    public class Win32
    {


        [DllImportAttribute("kernel32.dll")]
        public static extern int GetConsoleWindow();

       
        public Win32()
        {

        }

        ~Win32()
        {
        }
        
    }
}
