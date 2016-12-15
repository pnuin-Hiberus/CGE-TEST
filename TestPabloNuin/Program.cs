using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AprovisionamientoO365.Base;


namespace TestPabloNuin.Aprovisionamiento
{
    class Program
    {
        static void Main(string[] args)
        {
            AprovisionamientoBase.Empieza("SettingsAprovisionamiento.xml", "TestPabloNuin-Aprovisionamiento");
        }
    }
}
