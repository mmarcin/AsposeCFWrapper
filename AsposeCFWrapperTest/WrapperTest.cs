using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AsposeCF; 


namespace AsposeCFWrapperTest
{
    public static class Test
    {
        static void Main()
        {
            Wrapper AW = new Wrapper("c:/projects/eProcurement/www_root/prototypes/", "Template.docx", "c:/projects/eProcurement/www_root/prototypes/", "Out.Docx");
            AW.ExecuteRegions("SELECT ProcurementCode, ProcurementTitle FROM pro_tProcurement WHERE InstanceID = 9 ORDER BY ProcurementTitle", "Procurements");
            AW.ExecuteRegions("SELECT ContractCode, ContractTitle FROM cnt_tContract WHERE InstanceID = 18", "Contracts");
            AW.Save();
            Console.Write("Done.");
           
        }
    }
}
