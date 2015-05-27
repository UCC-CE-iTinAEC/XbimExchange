﻿using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xbim.COBieLiteUK;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Xbim.CobieLiteUK.Validation;
using Xbim.IO;
using XbimExchanger.IfcToCOBieLiteUK;
using Attribute = Xbim.COBieLiteUK.Attribute;
using System = Xbim.COBieLiteUK.System;


namespace Tests
{
    [TestClass]
    public class CoBieLiteUkTests
    {
        [TestMethod]
        public void CustomEnumerationTests()
        {
            var f = new Facility();
            Assert.AreEqual(AreaUnit.notdefined, f.AreaUnits);

            f.AreaUnits = AreaUnit.squarefeet;
            Assert.AreEqual(AreaUnit.squarefeet, f.AreaUnits);

            //this should be qualified as a user defined value
            f.AreaUnitsCustom = "aaa";
            Assert.AreEqual(AreaUnit.userdefined, f.AreaUnits);

            //this should be picked from aliases
            f.AreaUnitsCustom = "Square centimeters";
            Assert.AreEqual(AreaUnit.squarecentimeters, f.AreaUnits);

            //this shouldn't change the value
            f.AreaUnits = AreaUnit.userdefined;
            Assert.AreEqual(AreaUnit.squarecentimeters, f.AreaUnits);

            //this should set custom area units to null
            f.AreaUnits = AreaUnit.notdefined;
            Assert.IsNull(f.AreaUnitsCustom);
        }

        //[TestMethod]
        //public void IFSModelAnalyses()
        //{
        //    const string file = @"c:\CODE\SampleData\IFS\DB4 full Model A_DPoW.json";
        //    //**************** semantical analysis
        //    var f = Facility.ReadJson(file);

        //    var assemblies = f.Get<Assembly>().Where(i => i != null);
        //    var assets = f.Get<Asset>().Where(i => i != null);
        //    var assetsTypes = (f.AssetTypes ?? new List<AssetType>()).Where(i => i != null);
        //    var attributes = f.Get<Attribute>().Where(i => i != null);
        //    var connections = f.Get<Connection>().Where(i => i != null);
        //    var contacts = (f.Contacts ?? new List<Contact>()).Where(i => i != null);
        //    var documents = f.Get<Document>().Where(i => i != null);
        //    var floors = (f.Floors ?? new List<Floor>()).Where(i => i != null);
        //    var impacts = f.Get<Impact>().Where(i => i != null);
        //    var issues = f.Get<Issue>().Where(i => i != null);
        //    var jobs = f.Get<Job>().Where(i => i != null);
        //    var resources = f.Get<Resource>().Where(i => i != null);
        //    var spaces = f.Get<Space>().Where(i => i != null);
        //    var spares = f.Get<Spare>().Where(i => i != null);
        //    var systems = (f.Systems ?? new List<Xbim.COBieLiteUK.System>()).Where(i => i != null);
        //    var zones = (f.Zones ?? new List<Zone>()).Where(i => i != null);

        //    //report
        //    Debug.WriteLine("Assemblies: {0}", assemblies.Count());
        //    Debug.WriteLine("Assets: {0}", assets.Count());
        //    Debug.WriteLine("AssetTypes: {0}", assetsTypes.Count());
        //    Debug.WriteLine("Attributes: {0}", attributes.Count());
        //    Debug.WriteLine("Connections: {0}", connections.Count());
        //    Debug.WriteLine("Contacts: {0}", contacts.Count());
        //    Debug.WriteLine("Documents: {0}", documents.Count());
        //    Debug.WriteLine("Floors: {0}", floors.Count());
        //    Debug.WriteLine("Impacts: {0}", impacts.Count());
        //    Debug.WriteLine("Issues: {0}", issues.Count());
        //    Debug.WriteLine("Jobs: {0}", jobs.Count());
        //    Debug.WriteLine("Resources: {0}", resources.Count());
        //    Debug.WriteLine("Spaces: {0}", spaces.Count());
        //    Debug.WriteLine("Spares: {0}", spares.Count());
        //    Debug.WriteLine("Systems: {0}", systems.Count());
        //    Debug.WriteLine("Zones: {0}", zones.Count());
        //}

        //[TestMethod]
        //public void IFSFileAnalyses()
        //{
        //    const string file = @"c:\CODE\SampleData\IFS\DB4 full Model A_DPoW.json";
        //    //const string file = @"c:\Users\mxfm2\Desktop\SampleHouse.json";
        //    const string result = @"c:\CODE\SampleData\IFS\Tokens.txt";


        //    //**************** lexical analysis
        //    var tokens = new Dictionary<string, int>();
        //    var filesize = 0;
        //    using (var reader = File.OpenText(file))
        //    {
        //        var data = reader.ReadToEnd();
        //        filesize = data.Length;
        //        //analyse repetition of property names
        //        var propNameRegex = new Regex("(?<=\")([\\w0-9]+?)(?=\":)");
        //        var matches = propNameRegex.Matches(data).OfType<Match>();
        //        reader.Close();
        //        foreach (var match in matches)
        //        {
        //            if (!tokens.Keys.Contains(match.Value)) tokens.Add(match.Value, 0);
        //            tokens[match.Value]++;
        //        }
        //        data = null;
        //        matches = null;
        //    }

        //    using (var w = File.CreateText(result))
        //    {
        //        var total = tokens.Values.Aggregate(0, (a, b) => a + b);
        //        w.WriteLine("Total count of tokens: {0:## ### ###}", total);
        //        w.WriteLine("Total file length: {0:## ### ###}", filesize);
        //        w.WriteLine("------------------------------");
        //        foreach (var t in tokens.OrderByDescending(t => t.Value))
        //        {
        //            w.WriteLine("{0,30} {1,10} {2,10:P} {3,10:P}", t.Key, t.Value, (float) t.Value/(float) total,
        //                (float) (t.Key.Length*t.Value)/(float) filesize);
        //        }
        //    }
        //}

        [TestMethod]
        public void CoBieLiteUkCreation()
        {
            var facility = new Facility
            {
                CreatedOn = DateTime.Now,
                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                Categories =
                    new List<Category>
                    {
                        new Category {Code = "Bd_34_54", Description = "Schools", Classification = "Sample"}
                    },
                ExternalId = Guid.NewGuid().ToString(),
                AreaUnits = AreaUnit.squaremeters,
                CurrencyUnit = CurrencyUnit.GBP,
                LinearUnits = LinearUnit.millimeters,
                VolumeUnits = VolumeUnit.cubicmeters,
                AreaMeasurement = "NRM",
                Phase = "Phase A",
                Description = "New facility description",
                Name = "Ellison Building",
                Project = new Project
                {
                    ExternalId = Guid.NewGuid().ToString(),
                    Name = "Project A"
                },
                Site = new Site
                {
                    ExternalId = Guid.NewGuid().ToString(),
                    Name = "Site A"
                },
                Zones = new List<Zone>
                {
                    new Zone
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        ExternalId = Guid.NewGuid().ToString(),
                        Name = "Zone A",
                        Categories = new List<Category> {new Category {Code = "45.789.78", Classification = "Sample"}},
                        Description = "Description of the zone A",
                        Spaces = new List<SpaceKey>
                        {
                            new SpaceKey {Name = "A001 - Front Room"},
                            new SpaceKey {Name = "A002 - Living Room"},
                            new SpaceKey {Name = "A003 - Bedroom"},
                        }
                    }
                },
                Contacts = new List<Contact>
                {
                    new Contact
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Categories = new List<Category> {new Category {Code = "12.45.56", Classification = "Sample"}},
                        FamilyName = "Martin",
                        Email = "martin.cerny@northumbria.ac.uk",
                        GivenName = "Cerny"
                    },
                    new Contact
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Categories = new List<Category> {new Category {Code = "12.45.56", Classification = "Sample"}},
                        FamilyName = "Peter",
                        Email = "peter.pan@northumbria.ac.uk",
                        GivenName = "Pan"
                    },
                    new Contact
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Categories = new List<Category> {new Category {Code = "12.45.56", Classification = "Sample"}},
                        FamilyName = "Paul",
                        Email = "paul.mccartney@northumbria.ac.uk",
                        GivenName = "McCartney"
                    }
                },
                Floors = new List<Floor>
                {
                    new Floor
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Elevation = 15000,
                        Height = 3400,
                        Spaces = new List<Space>
                        {
                            new Space
                            {
                                CreatedOn = DateTime.Now,
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                Categories =
                                    new List<Category> {new Category {Code = "Sp_02_78_98", Classification = "Sample"}},
                                Description = "First front room in COBieLiteUK ever",
                                Name = "A001 - Front Room",
                                UsableHeight = 3500,
                                NetArea = 6
                            },
                            new Space
                            {
                                CreatedOn = DateTime.Now,
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                Categories =
                                    new List<Category> {new Category {Code = "Sp_02_78_98", Classification = "Sample"}},
                                Description = "First living room in COBieLiteUK ever",
                                Name = "A002 - Living Room",
                                UsableHeight = 4200,
                                NetArea = 55
                            },
                            new Space
                            {
                                CreatedOn = DateTime.Now,
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                Categories =
                                    new List<Category> {new Category {Code = "Sp_02_78_98", Classification = "Sample"}},
                                Description = "First bedroom in COBieLiteUK ever",
                                Name = "A003 - Bedroom",
                                UsableHeight = 4100,
                                NetArea = 25
                            }
                        }
                    }
                },
                AssetTypes = new List<AssetType>
                {
                    new AssetType
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Brick layered wall",
                        Assets = new List<Asset>
                        {
                            new Asset
                            {
                                CreatedOn = DateTime.Now,
                                Name = "120mm partition wall",
                                Representations = new List<Representation>
                                {
                                    new Representation
                                    {
                                        CreatedOn = DateTime.Now,
                                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                        X = 0,
                                        Y = 0,
                                        Z = 0,
                                        SizeX = 1000,
                                        SizeY = 2000,
                                        SizeZ = 200,
                                        Name = Guid.NewGuid().ToString()
                                    }
                                },
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                            },
                            new Asset
                            {
                                CreatedOn = DateTime.Now,
                                Name = "180mm partition wall",
                                Representations = new List<Representation>
                                {
                                    new Representation
                                    {
                                        CreatedOn = DateTime.Now,
                                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                        X = 0,
                                        Y = 0,
                                        Z = 0,
                                        SizeX = 1000,
                                        SizeY = 2000,
                                        SizeZ = 200,
                                        Name = Guid.NewGuid().ToString()
                                    }
                                },
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                            },
                            new Asset
                            {
                                CreatedOn = DateTime.Now,
                                Name = "350mm external brick wall",
                                Representations = new List<Representation>
                                {
                                    new Representation
                                    {
                                        CreatedOn = DateTime.Now,
                                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                                        X = 0,
                                        Y = 0,
                                        Z = 0,
                                        SizeX = 1000,
                                        SizeY = 2000,
                                        SizeZ = 200,
                                        Name = Guid.NewGuid().ToString()
                                    }
                                },
                                CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                            }
                        }
                    }
                },
                Attributes = new List<Attribute>
                {
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "String attribute",
                        Value = new StringAttributeValue {Value = "Almukantarant"},
                        Categories = new List<Category> {new Category {Code = "Submitted", Classification = "Sample"}},
                    },
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Boolean attribute",
                        Value = new BooleanAttributeValue {Value = true},
                        Categories = new List<Category> {new Category {Code = "Submitted", Classification = "Sample"}},
                    },
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Datetime attribute",
                        Value = new DateTimeAttributeValue {Value = DateTime.Now},
                        Categories = new List<Category> {new Category {Code = "Submitted", Classification = "Sample"}},
                    },
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Decimal attribute",
                        Value = new DecimalAttributeValue {Value = 256.2},
                        Categories = new List<Category> {new Category {Code = "Submitted", Classification = "Sample"}},
                    },
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Integer attribute",
                        Value = new IntegerAttributeValue {Value = 7},
                        Categories = new List<Category> {new Category {Code = "Submitted", Classification = "Sample"}},
                    },
                    new Attribute
                    {
                        CreatedOn = DateTime.Now,
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"},
                        Name = "Null attribute"
                    }
                },
                Stages = new List<ProjectStage>(new[]
                {
                    new ProjectStage
                    {
                        Name = "Stage 0",
                        CreatedOn = DateTime.Now,
                        Start = DateTime.Now.AddDays(5),
                        End = DateTime.Now.AddDays(10),
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                    },
                    new ProjectStage
                    {
                        Name = "Stage 1",
                        CreatedOn = DateTime.Now,
                        Start = DateTime.Now.AddDays(10),
                        End = DateTime.Now.AddDays(20),
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                    },
                    new ProjectStage
                    {
                        Name = "Stage 2",
                        CreatedOn = DateTime.Now,
                        Start = DateTime.Now.AddDays(20),
                        End = DateTime.Now.AddDays(110),
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                    },
                    new ProjectStage
                    {
                        Name = "Stage 3",
                        CreatedOn = DateTime.Now,
                        Start = DateTime.Now.AddDays(110),
                        End = DateTime.Now.AddDays(300),
                        CreatedBy = new ContactKey {Email = "martin.cerny@northumbria.ac.uk"}
                    },
                })
            };

            //save model to file to check it
            string msg;
            const string xmlFile = "facility.cobielite.xml";
            const string jsonFile = "facility.cobielite.json";
            const string xlsxFile = "facility.cobielite.xlsx";
            facility.WriteXml(xmlFile, true);
            facility.WriteJson(jsonFile, true);
            facility.WriteCobie(xlsxFile, out msg);

            var facility2 = Facility.ReadXml(xmlFile);
            var facility3 = Facility.ReadJson(jsonFile);
        }

        [TestMethod]
        public void AttributeTest()
        {
            var facility = new Facility
            {
                Attributes = new List<Attribute>
                {
                    new Attribute
                    {
                        Name = "Boolean",
                        Value = new BooleanAttributeValue {Value = true}
                    },
                    new Attribute
                    {
                        Name = "DateTime",
                        Value = new DateTimeAttributeValue {Value = DateTime.Today}
                    },
                    new Attribute
                    {
                        Name = "Decimal",
                        Value = new DecimalAttributeValue {Value = 10.0}
                    },
                    new Attribute
                    {
                        Name = "Integer",
                        Value = new IntegerAttributeValue {Value = 5}
                    },
                    new Attribute
                    {
                        Name = "String",
                        Value = new StringAttributeValue {Value = "A"}
                    },
                    new Attribute
                    {
                        Name = "Null"
                    }
                }
            };

            const string file = "attribute_test.json";

            //write to file and minified file (attribute type names minified)
            facility.WriteJson(file);

            var f2 = Facility.ReadJson(file);

            foreach (var f in new []{f2})
            {
                Assert.AreEqual((f.Attributes.FirstOrDefault(a => a.Name == "Boolean").Value as BooleanAttributeValue).Value, true);
                Assert.AreEqual((f.Attributes.FirstOrDefault(a => a.Name == "DateTime").Value as DateTimeAttributeValue).Value, DateTime.Today);
                Assert.AreEqual((f.Attributes.FirstOrDefault(a => a.Name == "Decimal").Value as DecimalAttributeValue).Value ?? 0, 10.0, 0.0000001);
                Assert.AreEqual((f.Attributes.FirstOrDefault(a => a.Name == "Integer").Value as IntegerAttributeValue).Value, 5);
                Assert.AreEqual((f.Attributes.FirstOrDefault(a => a.Name == "String").Value as StringAttributeValue).Value, "A");
                Assert.IsNull(f.Attributes.FirstOrDefault(a => a.Name == "Null").Value);
            }

        }

        [TestMethod]
        [DeploymentItem("TestFiles\\2012-03-23-Duplex-Design.xlsx")]
        public void ReadingSpreadsheet()
        {
            string msg;
            var facility = Facility.ReadCobie("2012-03-23-Duplex-Design.xlsx", out msg);
            facility.WriteJson("..\\..\\2012-03-23-Duplex-Design.cobielite.json", true);

            Assert.AreEqual(AreaUnit.squaremeters, facility.AreaUnits);
            Assert.IsFalse(String.IsNullOrEmpty(msg));

            var log = new StringWriter();
            facility.ValidateUK2012(log, true);
            Debug.Write(log.ToString());
            Debug.WriteLine("----------------------------------------------------------------------");

            //second run after fixings
            log = new StringWriter();
            facility.ValidateUK2012(log, true);
            Debug.Write(log.ToString());

            facility.WriteCobie("..\\..\\2012-03-23-Duplex-Design.fixed.xlsx", out msg);

            var f2 = Facility.ReadJson("..\\..\\2012-03-23-Duplex-Design.cobielite.json");
        }

        [TestMethod]
        [DeploymentItem("TestFiles\\OBN1-COBie-UK-2014.xlsx")]
        public void ReadingUkSpreadsheet()
        {
            string msg;
            var facility = Facility.ReadCobie("OBN1-COBie-UK-2014.xlsx", out msg);
            facility.WriteJson("..\\..\\OBN1-COBie-UK-2014.cobielite.json", true);

            var log = new StringWriter();
            facility.ValidateUK2012(log, true);
            Debug.Write(log.ToString());
        }

        //[TestMethod]
        public void CobieFix()
        {
            var files = new[]
            {
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\Ext01.xlsx",
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\Ext01.fixed.xlsx",
                @"C:\Users\mxfm2\Downloads\Bad Cobie\Ext01.xls",
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\Struc.xls",
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\Site.xls",
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\INT02.xls",
                //@"C:\Users\mxfm2\Downloads\Bad Cobie\Int01.xls"
            };
            foreach (var file in files)
            {
                Stopwatch completeWatch = new Stopwatch();
                completeWatch.Start();

                var dir = Path.GetDirectoryName(file);
                var name = Path.GetFileNameWithoutExtension(file);
                var newFile = Path.Combine(dir ?? "", name + ".fixed.xlsx");
                Debug.WriteLine("============ Processing: " + (name ?? ""));

                using (var log = File.CreateText(Path.Combine(dir ?? "", name + ".fixed.txt")))
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start();
                    string msg;
                    var facility = Facility.ReadCobie(file, out msg);
                    stopWatch.Stop();

                    if (!String.IsNullOrEmpty(msg))
                        log.WriteLine(msg);
                    Debug.WriteLine("Reading COBie: " + stopWatch.ElapsedMilliseconds);

                    stopWatch.Reset();
                    stopWatch.Start();
                    facility.ValidateUK2012(log, true);
                    stopWatch.Stop();

                    Debug.WriteLine("Validating COBie: " + stopWatch.ElapsedMilliseconds);


                    stopWatch.Reset();
                    stopWatch.Start();
                    //Debug.Write(msg);
                    //Debug.Write(log.ToString());    
                    facility.WriteCobie(newFile, out msg);
                    stopWatch.Stop();
                    if (!String.IsNullOrEmpty(msg))
                        log.WriteLine(msg);
                    Debug.WriteLine("Writing COBie: " + stopWatch.ElapsedMilliseconds);
                    log.Close();
                }

                completeWatch.Stop();
                Debug.WriteLine("========== Complete processing of {0}: {1}ms", name, completeWatch.ElapsedMilliseconds);
            }
        }

        [TestMethod]
        public void CobieAttributesCreation()
        {
            int i = 1;
            var attI = AttributeValue.CreateFromObject(i);

            Int16 i16 = 13;
            attI = AttributeValue.CreateFromObject(i16);

            Int32 i32 = 13;
            attI = AttributeValue.CreateFromObject(i32);

            DateTime d = DateTime.Now;
            var attD = AttributeValue.CreateFromObject(i);

            string s = "Yes";
            var attS = AttributeValue.CreateFromObject(s);

            bool b = true;
            var attB = AttributeValue.CreateFromObject(b);

            double dbl = 3.14;
            var attDbl = AttributeValue.CreateFromObject(dbl);

            AttributeValue dAt = new DecimalAttributeValue() {Value = Math.E};
            var fromA = AttributeValue.CreateFromObject(dAt);
        }

        [TestMethod]
        [DeploymentItem("TestFiles\\2012-03-23-Duplex-Design.xlsx")]
        public void WritingSpreadsheet()
        {
            string msg;
            var facility = Facility.ReadCobie("2012-03-23-Duplex-Design.xlsx", out msg);
            facility.WriteCobie("..\\..\\2012-03-23-Duplex-Design_enhanced.xlsx", out msg);
        }

        [TestMethod]
        [DeploymentItem("TestFiles\\OBN1-COBie-UK-2014.xlsx")]
        public void WritingUkSpreadsheet()
        {
            string msg;
            var facility = Facility.ReadCobie("OBN1-COBie-UK-2014.xlsx", out msg);
            facility.WriteCobie("..\\..\\OBN1-COBie-UK-2014_plain.xlsx", out msg, "UK2012", false);
        }

        [TestMethod]
        public void DeepSearchTest()
        {
            #region Model

            var facility = new Facility
            {
                Contacts = new List<Contact>(new[]
                {
                    new Contact
                    {
                        Name = "martin.cerny@northumbria.ac.uk"
                    }
                }),
                Floors = new List<Floor>(new[]
                {
                    new Floor
                    {
                        Name = "Floor 0",
                        Spaces = new List<Space>(new[]
                        {
                            new Space
                            {
                                Name = "Space A",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space A attribute 1"},
                                    new Attribute {Name = "Space A attribute 2"},
                                })
                            },
                            new Space
                            {
                                Name = "Space B",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space B attribute 1"},
                                    new Attribute {Name = "Space B attribute 2"},
                                })
                            }
                        })
                    },
                    new Floor
                    {
                        Name = "Floor 1",
                        Spaces = new List<Space>(new[]
                        {
                            new Space
                            {
                                Name = "Space C",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space C attribute 1"},
                                    new Attribute {Name = "Space C attribute 2"},
                                })
                            },
                            new Space
                            {
                                Name = "Space D",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space D attribute 1"},
                                    new Attribute {Name = "Space D attribute 2"},
                                })
                            }
                        })
                    },
                    new Floor
                    {
                        Name = "Floor 2",
                        Spaces = new List<Space>(new[]
                        {
                            new Space
                            {
                                Name = "Space E",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space E attribute 1"},
                                    new Attribute {Name = "Space E attribute 2"},
                                })
                            },
                            new Space
                            {
                                Name = "Space F",
                                Attributes = new List<Attribute>(new[]
                                {
                                    new Attribute {Name = "Space F attribute 1"},
                                    new Attribute {Name = "Space F attribute 2"},
                                })
                            }
                        })
                    }
                })
            };

            #endregion

            var allAttributes = facility.Get<Attribute>();
            Assert.AreEqual(12, allAttributes.Count());

            var allSpaces = facility.Get<Space>();
            Assert.AreEqual(6, allSpaces.Count());

            var spaceA = facility.Get<Space>(s => s.Name == "Space A");
            Assert.AreEqual(1, spaceA.Count());

            var self = facility.Get<Facility>().FirstOrDefault();
            Assert.IsNotNull(self);

            var contact =
                facility.Get<CobieObject>(
                    c => c.GetType() == typeof (Contact) && c.Name == "martin.cerny@northumbria.ac.uk");
        }

        [DeploymentItem("TestFiles\\OBN1-COBie-UK-2014.xlsx")]
        [TestMethod]
        [DeploymentItem("ValidationFiles\\Lakeside_Restaurant.ifc")]
        [DeploymentItem("RIBAETestFiles\\001-Kenton_High_School_Model.ifc")]
        [DeploymentItem("RIBAETestFiles\\001 Hello Wall.ifc")]
        [DeploymentItem("RIBAETestFiles\\Duplex_A_20110907_optimized.ifc")]
        [DeploymentItem("RIBAETestFiles\\NBS_LakesideRestaurant_small_optimized.ifc")]
        [DeploymentItem("RIBAETestFiles\\Office_A_20110811_optimized.ifc")]
        public void IfcToCoBieLiteUkTest()
        {
            string[] testFiles = new string[] { "Lakeside_Restaurant.ifc", "001-Kenton_High_School_Model.ifc", "001 Hello Wall.ifc", "Duplex_A_20110907_optimized.ifc", "NBS_LakesideRestaurant_small_optimized.ifc", "Office_A_20110811_optimized.ifc" };

            foreach (var ifcTestFile in testFiles)
            {
                using (var m = new XbimModel())
                {
                    var xbimTestFile = Path.ChangeExtension(ifcTestFile, "xbim");
                    var jsonFile = Path.ChangeExtension(ifcTestFile, "json");
                    m.CreateFrom(ifcTestFile, xbimTestFile, null, true, true);
                    var facilities = new List<Facility>();
                    var ifcToCoBieLiteUkExchanger = new IfcToCOBieLiteUkExchanger(m, facilities);
                    facilities = ifcToCoBieLiteUkExchanger.Convert();

                    foreach (var facilityType in facilities)
                    {
                        var log = new StringWriter();
                        facilityType.ValidateUK2012(log, true);

                        string msg;
                        facilityType.WriteJson(jsonFile, true);
                        facilityType.WriteCobie("..\\..\\" + System.IO.Path.ChangeExtension(ifcTestFile, ".xlsx"), out msg, "UK2012", true);


                        break;
                    }
                }
            }
        }
        [TestMethod]
        [DeploymentItem("RIBAETestFiles\\001 BTK Sample.ifc")]
        public void IfcToCoBieLiteUkTestSingleFile()
        {
            string ifcTestFile = "001 BTK Sample.ifc";

            using (var m = new XbimModel())
            {
                var xbimTestFile = Path.ChangeExtension(ifcTestFile, "xbim");
                var jsonFile = Path.ChangeExtension(ifcTestFile, "json");
                m.CreateFrom(ifcTestFile, xbimTestFile, null, true, true);
                var facilities = new List<Facility>();
                var ifcToCoBieLiteUkExchanger = new IfcToCOBieLiteUkExchanger(m, facilities);
                facilities = ifcToCoBieLiteUkExchanger.Convert();

                foreach (var facilityType in facilities)
                {
                    var log = new StringWriter();
                    facilityType.ValidateUK2012(log, true);

                    string msg;
                    facilityType.WriteJson(jsonFile, true);
                    facilityType.WriteCobie("..\\..\\" + System.IO.Path.ChangeExtension(ifcTestFile, ".xlsx"), out msg, "UK2012", true);


                    break;
                }
            }
        }

        [DeploymentItem("ValidationFiles\\Lakeside_Restaurant.json")]
        [DeploymentItem("ValidationFiles\\Lakeside_Restaurant-stage6-COBie.json")]
        [TestMethod]
        public void RemoveUnrequiredAssetsFromSubmission()
        {
            var submitted = Facility.ReadJson("Lakeside_restaurant.json");
            Assert.IsNotNull(submitted.AssetTypes);
            var requirement = Facility.ReadJson("Lakeside_Restaurant-stage6-COBie.json");
            Assert.IsNotNull(requirement.AssetTypes);
            var submittedAssetTypes = new List<AssetType>();
            foreach (var assetTypeRequirement in requirement.AssetTypes)
            {
                var v = new CobieObjectValidator<AssetType, Asset>(assetTypeRequirement)
                {
                    TerminationMode = TerminationMode.StopOnFirstFail
                };
                var candidates = v.GetCandidates(submitted.AssetTypes).ToList();
                submittedAssetTypes.AddRange(candidates.Select(c => c.MatchedObject).Cast<AssetType>());
            }
            Assert.IsTrue(submittedAssetTypes.Count > 0);
            submitted.AssetTypes = submittedAssetTypes;
            string msg;
            submitted.WriteCobie("Lakeside_restaurant_submission.xlsx", out msg);
            Assert.IsTrue(string.IsNullOrWhiteSpace(msg));
        }
    }
}