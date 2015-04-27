using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xbim.CobieLiteUK.Validation;
using Xbim.CobieLiteUK.Validation.Reporting;
using Xbim.COBieLiteUK;
using Xbim.IO;
using XbimExchanger.IfcToCOBieLiteUK;
using System;

namespace Tests
{
    [TestClass]
    [DeploymentItem(@"ValidationFiles\")]
    [DeploymentItem(@"RIBAETestFiles\")]
    public class CobieLiteUKValidationTests
    {
        [TestMethod]
        public void CanSaveValidatedVacility()
        {
            var validated = GetValidated(@"Lakeside_Restaurant-stage6-COBie.json");
            validated.WriteJson(@"..\..\ValidationReport.json", true);
            validated.WriteXml(@"..\..\ValidationReport.xml", true);
            validated.WriteJson(@"ValidationReport.json", true);
        }

        [TestMethod]
        public void CanSaveValidationReport()
        {
            // stage 0 is for documents
            // stage 1 is for zones
            // stage 6 is for assettypes
            var validated = GetValidated(@"Lakeside_Restaurant-stage6-COBie.json");
            const string repName = @"..\..\ValidationReport.xlsx";
            var xRep = new ExcelValidationReport();
            var ret = xRep.Create(validated, repName);
            Assert.IsTrue(ret, "File not created");
        }

        private static Facility GetValidated(string requirementFile)
        {
            const string ifcTestFile = @"Lakeside_Restaurant_fabric_only.ifczip";
            Facility sub = null;

            //create validation file from IFC
            using (var m = new XbimModel())
            {
                var xbimTestFile = Path.ChangeExtension(ifcTestFile, "xbim");
                m.CreateFrom(ifcTestFile, xbimTestFile, null, true, true);
                var facilities = new List<Facility>();
                var ifcToCoBieLiteUkExchanger = new IfcToCOBieLiteUkExchanger(m, facilities);
                facilities = ifcToCoBieLiteUkExchanger.Convert();
                sub = facilities.FirstOrDefault();
            }
            Assert.IsTrue(sub != null);
            var vd = new FacilityValidator();
            var req = Facility.ReadJson(requirementFile);
            var validated = vd.Validate(req, sub);
            return validated;
        }

        [TestMethod]
        public void ValidateXlsLakeside()
        {
            const string xlsx = @"LakesideWithDocuments.xls";
            string msg;
            var cobie = Facility.ReadCobie(xlsx, out msg);
            var req = Facility.ReadJson(@"Lakeside_Restaurant-stage6-COBie.json");
            var validator = new FacilityValidator();
            var result = validator.Validate(req, cobie);
            result.WriteJson(@"..\..\XlsLakesideWithDocumentsValidationStage6.json", true);
        }

        [TestMethod]
        public void ValidateXlsLakeside2()
        {
            const string xlsx = @"c:\Users\mxfm2\Dropbox\Martin\Lakeside_Restaurant_fabric_only.xlsx";
            string msg;
            var cobie = Facility.ReadCobie(xlsx, out msg);
            var req = Facility.ReadJson(@"c:\Users\mxfm2\Dropbox\Martin\003-Lakeside_Restaurant-stage6-COBie.json");
            var validator = new FacilityValidator();
            var result = validator.Validate(req, cobie);

            //create report
            using (var stream = File.Create(@"c:\Users\mxfm2\Dropbox\Martin\Lakeside_Restaurant_fabric_only.report.xlsx"))
            {
                var report = new ExcelValidationReport();
                report.Create(result, stream, ExcelValidationReport.SpreadSheetFormat.Xlsx);
                stream.Close();
            }
        }

        [TestMethod]
        public void ValidateXlsLakesideForStage0()
        {
            var result = LakeSide0();
            result.WriteJson(@"..\..\XlsLakesideWithDocumentsValidationStage0.json", true);
        }

        [TestMethod]
        public void LakeSideXls0ValidationReport()
        {
            var validated = LakeSide0();
            const string repName = @"..\..\LakeSideXls0ValidationReport.xlsx";
            var xRep = new ExcelValidationReport();
            var ret = xRep.Create(validated, repName);
            Assert.IsTrue(ret, "File not created");
        }

        private static Facility LakeSide0()
        {
            const string xlsx = @"LakesideWithDocuments.xls";
            string msg;
            var cobie = Facility.ReadCobie(xlsx, out msg);
            var req = Facility.ReadJson(@"Lakeside_Restaurant-stage0-COBie.json");
            var validator = new FacilityValidator();
            var result = validator.Validate(req, cobie);
            return result;
        }

        [TestMethod]
        public void VerifyCobie()
        {
            string reportBlobName = string.Format("{0}{1}", "output", ".json");
            string ext = ".ifc";
            string ext2 = ".json";

            Stream input = File.OpenRead("NBS_LakesideRestaurant_small_optimized.ifc");
            Stream inputRequirements = File.OpenRead("003-Lakeside_Restaurant-stage6-COBie.json"); 
            
            Facility facility = null;
            string msg;
            switch (ext)
            {
                case ".ifc":
                case ".ifczip":
                case ".ifcxml":
                    facility = GetFacilityFromIfc(input, ".ifc");
                    break;
                case ".json":
                    facility = Facility.ReadJson(input);
                    break;
                case ".xml":
                    facility = Facility.ReadXml(input);
                    break;
                case ".xls":
                    facility = Facility.ReadCobie(input, ExcelTypeEnum.XLS, out msg);
                    break;
                case ".xlsx":
                    facility = Facility.ReadCobie(input, ExcelTypeEnum.XLSX, out msg);
                    break;
            }


            Facility requirements = null;
            switch (ext2)
            {
                case ".xml":
                    requirements = Facility.ReadXml(inputRequirements);
                    break;
                case ".json":
                    requirements = Facility.ReadJson(inputRequirements);
                    break;
                case ".xls":
                    requirements = Facility.ReadCobie(inputRequirements, ExcelTypeEnum.XLS, out msg);
                    break;
                case ".xlsx":
                    requirements = Facility.ReadCobie(inputRequirements, ExcelTypeEnum.XLSX, out msg);
                    break;
            }

            if (facility == null || requirements == null)
                return;

            var vd = new FacilityValidator();
            var validated = vd.Validate(requirements, facility);
            using (var repStream = File.OpenWrite(reportBlobName))
            {
                validated.WriteJson(repStream);
                repStream.Close();
            }

            var rep = new ExcelValidationReport();
            var temp = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetTempFileName(), ".xlsx"));

            try
            {
                rep.Create(facility, temp, ExcelValidationReport.SpreadSheetFormat.Xlsx);
            }
            finally
            {
                if (File.Exists(temp)) File.Delete(temp);
            }

        }

        private static Facility GetFacilityFromIfc(Stream file, string extension)
        {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try
            {
                //store temporarily
                using (var fileStream = File.OpenWrite(temp))
                {
                    file.CopyTo(fileStream);
                    fileStream.Flush();
                    fileStream.Close();
                }

                using (var model = new XbimModel())
                {
                    model.CreateFrom(temp, null, null, true);

                    var facilities = new List<Facility>();
                    var ifcToCoBieLiteUkExchanger = new IfcToCOBieLiteUkExchanger(model, facilities);

                    return ifcToCoBieLiteUkExchanger.Convert().FirstOrDefault();
                }
            }
            //tidy up
            finally
            {
                if (File.Exists(temp)) File.Delete(temp);
            }


        }
    }
}
