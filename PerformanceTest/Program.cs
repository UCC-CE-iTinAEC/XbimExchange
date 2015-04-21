using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.DPoW;
using Xbim.IO;
using Xbim.XbimExtensions.Interfaces;
using XbimExchanger.DPoWToCOBieLiteUK;

namespace PerformanceTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var pow = PlanOfWork.OpenJson("013-Lakeside_Restaurant.dpow");
            const string dir = "..\\..\\COBieLiteUK";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            string msg;
            foreach (var stage in pow.ProjectStages)
            {
                var json = Path.Combine(dir, stage.Name + ".cobie.json");
                var xlsx = Path.Combine(dir, stage.Name + ".xlsx");
                var ifc = Path.Combine(dir, stage.Name + ".ifc");

                var facility = new Xbim.COBieLiteUK.Facility();
                var cobieExchanger = new DPoWToCOBieLiteUKExchanger(pow, facility, stage);
                cobieExchanger.Convert();

                facility.WriteJson(json, true);
                facility.WriteCobie(xlsx, out msg);


                using (var ifcModel = XbimModel.CreateTemporaryModel())
                {
                    ifcModel.Initialise("Xbim Tester", "XbimTeam", "Xbim.Exchanger", "Xbim Development Team", "3.0");
                    ifcModel.Header.FileName.Name = stage.Name;
                    ifcModel.ReloadModelFactors();
                    using (var txn = ifcModel.BeginTransaction("Conversion from COBie"))
                    {
                        var ifcExchanger = new XbimExchanger.COBieLiteUkToIfc.CoBieLiteUkToIfcExchanger(facility, ifcModel);
                        ifcExchanger.Convert();
                        txn.Commit();
                    }
                    ifcModel.SaveAs(ifc, XbimStorageType.IFC);
                    ifcModel.Close();
                }
            }
        }
    }
}
