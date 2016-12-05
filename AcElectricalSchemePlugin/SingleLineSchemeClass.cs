using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;
using System.Data;
using System.Data.OleDb;

namespace AcElectricalSchemePlugin
{
    static class SingleLineSchemeClass
    {
        private static Document acDoc;
        private static Editor editor;
        private static Point3d currentPoint;
        private static Point3d prevAutomatic;
        private static bool firstUnit = true;
        private static int sinFilter = 1;
        private static int currentSection = 1;

        private struct Unit
        {
            public int automatic; //33

            public string phases; //2
            public string planName; //3
            public string nominalPower; //4
            public string calcPower; //5
            public string calcElectriciry; //6
            public string consumerName; //7
            public string contructName; //8
            public string cosPhi; //9
            public string calcElecFold; //10
            public string allowableDeviations; //11
            public string cableCrossSection; //12

            public string switchName; //13
            public string switchType; //14
            public string switchNominalElec; //15
            public string switchNominalVoltage; //16

            public string protectName; //17
            public string protectType; //18
            public string protectNominalElec; //19
            public string protectNominalVoltage; //20
            public string disconnector; //21
            public string protectTermDefense; //22
            public string shortCircuitDefense; //23
            public string protectDefenseReactTime; //24
            public string ultBreakCapacity;//25
           
            public string switchboardName; //26
            public string switchboardType; //27
            public string switchboardNominalElec; //28
            public string switchboardNominalVoltage; //29
            public string switchboardTermDefense; //30
            public string switchboardDefenseReactTime; //31
        }

        public static void drawScheme(string fileName)
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            List<Unit> units = loadData(fileName);

            PromptPointResult startPoint = editor.GetPoint("Выберите левый верхний угол первого листа в чертеже");
            if (startPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else currentPoint = startPoint.Value.Add(new Vector3d(110, 0, 0));

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    prevAutomatic = new Point3d();
                    firstUnit = true;
                    sinFilter = 1;
                    currentSection = 1;

                    for (int i = 0; i < units.Count; i++)
                    {
                        if (units[i].phases=="-")
                        {
                            currentSection++;
                            sinFilter = 1;
                        }
                        else insertUnit(acTrans, acModSpace, acDb, units[i]);
                    }
                    acTrans.Commit();
                }
            }
        }

        private static List<Unit> loadData(string path)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO;'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.FillSchema(dataSet, SchemaType.Mapped);
            adapter.Fill(dataSet);
            connection.Close();

            for (int column = 3; column < dataSet.Tables[0].Columns.Count; column++)
            {
                Unit unit = new Unit();
                unit.phases = dataSet.Tables[0].Rows[1][column].ToString();
                unit.planName = dataSet.Tables[0].Rows[2][column].ToString();
                unit.nominalPower = dataSet.Tables[0].Rows[3][column].ToString();
                unit.calcPower = dataSet.Tables[0].Rows[4][column].ToString();
                unit.calcElectriciry = dataSet.Tables[0].Rows[5][column].ToString();
                unit.consumerName = dataSet.Tables[0].Rows[6][column].ToString();
                unit.contructName = dataSet.Tables[0].Rows[7][column].ToString();
                unit.cosPhi = dataSet.Tables[0].Rows[8][column].ToString();
                unit.calcElecFold = dataSet.Tables[0].Rows[9][column].ToString();
                unit.allowableDeviations = dataSet.Tables[0].Rows[10][column].ToString();
                unit.cableCrossSection = dataSet.Tables[0].Rows[11][column].ToString();
                unit.switchName = dataSet.Tables[0].Rows[12][column].ToString();
                unit.switchType = dataSet.Tables[0].Rows[13][column].ToString();
                unit.switchNominalElec = dataSet.Tables[0].Rows[14][column].ToString();
                unit.switchNominalVoltage = dataSet.Tables[0].Rows[15][column].ToString();
                unit.protectName = dataSet.Tables[0].Rows[16][column].ToString();
                unit.protectType = dataSet.Tables[0].Rows[17][column].ToString();
                unit.protectNominalElec = dataSet.Tables[0].Rows[18][column].ToString();
                unit.protectNominalVoltage = dataSet.Tables[0].Rows[19][column].ToString();
                unit.disconnector = dataSet.Tables[0].Rows[20][column].ToString();
                unit.protectTermDefense = dataSet.Tables[0].Rows[21][column].ToString();
                unit.shortCircuitDefense = dataSet.Tables[0].Rows[22][column].ToString();
                unit.protectDefenseReactTime = dataSet.Tables[0].Rows[23][column].ToString();
                unit.ultBreakCapacity = dataSet.Tables[0].Rows[24][column].ToString();
                unit.switchboardName = dataSet.Tables[0].Rows[25][column].ToString();
                unit.switchboardType = dataSet.Tables[0].Rows[26][column].ToString();
                unit.switchboardNominalElec = dataSet.Tables[0].Rows[27][column].ToString();
                unit.switchboardNominalVoltage = dataSet.Tables[0].Rows[28][column].ToString();
                unit.switchboardTermDefense = dataSet.Tables[0].Rows[29][column].ToString();
                unit.switchboardDefenseReactTime = dataSet.Tables[0].Rows[30][column].ToString();
                string str = dataSet.Tables[0].Rows[32][column].ToString();
                if (str != "")
                {
                    unit.automatic = int.Parse(str);
                    units.Add(unit);
                }
                else break;
            }

            return units;
        }

        private static void insertUnit(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Unit unit)
        {
            Point3d minPoint = new Point3d(0, 0, 0);
            Point3d maxPoint = new Point3d(0, 0, 0);
            Point3d lowestPoint = new Point3d(0,0,0);
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\Automatic.dwg";
            string blockName = "Automatic";
            using (Database sourceDb = new Database(false, true))
            {
                if (System.IO.File.Exists(filename))
                {
                    sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                    using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                        if (bt.Has(blockName))
                            ids.Add(bt[blockName]);
                        trans.Commit();
                    }
                }
                else editor.WriteMessage("Не найден файл {0}", filename);
                if (ids.Count > 0)
                {
                    acTrans.TransactionManager.QueueForGraphicsFlush();
                    IdMapping iMap = new IdMapping();
                    acdb.WblockCloneObjects(ids, acdb.CurrentSpaceId, iMap, DuplicateRecordCloning.Replace, false);
                    BlockTable bt = (BlockTable)acTrans.GetObject(acdb.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(blockName))
                    {
                        BlockReference br = new BlockReference(currentPoint, bt[blockName]);
                        br.Layer = "0";
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);
                        if (br.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props =
                                br.DynamicBlockReferencePropertyCollection;

                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Видимость1" && !prop.ReadOnly)
                                {
                                    prop.Value = values[unit.automatic];
                                    break;
                                }
                            }
                        }
                        lowestPoint = currentPoint;
                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForWrite);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant) && attDef.Visible == true)
                            {
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "?ГР1KA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "?ГР1УЗО":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "?ГР2KM":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "?ГРКК1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "?ГРКМ1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "№ГР1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "№ГР1КМ":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardName;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_KM2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_КК1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_КМ":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_КМ1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.switchboardNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_УЗО":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "НОМИНАЛ_УЗО1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectNominalElec;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ТОК_УТЕЧКИ":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.ultBreakCapacity;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ТОК_УТЕЧКИ1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.ultBreakCapacity;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "УСТАВКА":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectTermDefense;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "УСТАВКА1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectTermDefense;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ФАЗА":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.phases.Contains("/") ? unit.phases.Split('/')[1] : "-";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ФАЗА1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.phases.Contains("/") ? unit.phases.Split('/')[1] : "-";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ФАЗА2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.phases.Contains("/") ? unit.phases.Split('/')[1] : "-";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "ХАРАКТЕРИСТИКА1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = unit.protectDefenseReactTime.Replace("Характеристика ", "");
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                                #endregion
                            }
                        }
                        getSize(br, ref minPoint, ref maxPoint);
                        lowestPoint = getLowestPoint(br);
                        if (firstUnit) { prevAutomatic = minPoint; firstUnit = false; }
                        while (minPoint.X < prevAutomatic.X)
                        {
                            br.Position = br.Position.Add(new Vector3d(1, 0, 0));
                            lowestPoint = lowestPoint.Add(new Vector3d(1, 0, 0));
                            minPoint = minPoint.Add(new Vector3d(1, 0, 0));
                            maxPoint = maxPoint.Add(new Vector3d(1, 0, 0));
                        }
                        prevAutomatic = maxPoint;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }

            if (!unit.consumerName.Contains("Резерв"))
            {
                if (unit.switchboardName.Contains("VFD"))
                {
                    ids = new ObjectIdCollection();
                    filename = @"Data\FreqConv.dwg";
                    blockName = "FreqConv";
                    using (Database sourceDb = new Database(false, true))
                    {
                        if (System.IO.File.Exists(filename))
                        {
                            sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                            using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                            {
                                BlockTable bt = (BlockTable)trans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                                if (bt.Has(blockName))
                                    ids.Add(bt[blockName]);
                                trans.Commit();
                            }
                        }
                        else editor.WriteMessage("Не найден файл {0}", filename);
                        if (ids.Count > 0)
                        {
                            acTrans.TransactionManager.QueueForGraphicsFlush();
                            IdMapping iMap = new IdMapping();
                            acdb.WblockCloneObjects(ids, acdb.CurrentSpaceId, iMap, DuplicateRecordCloning.Replace, false);
                            BlockTable bt = (BlockTable)acTrans.GetObject(acdb.BlockTableId, OpenMode.ForRead);
                            if (bt.Has(blockName))
                            {
                                BlockReference br = new BlockReference(lowestPoint, bt[blockName]);
                                br.Layer = "0";
                                modSpace.AppendEntity(br);
                                acTrans.AddNewlyCreatedDBObject(br, true);
                                BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                                foreach (ObjectId id in btr)
                                {
                                    DBObject obj = id.GetObject(OpenMode.ForWrite);
                                    AttributeDefinition attDef = obj as AttributeDefinition;
                                    if ((attDef != null) && (!attDef.Constant) && attDef.Visible == true)
                                    {
                                        #region attributes
                                        switch (attDef.Tag)
                                        {
                                            case "UZ":
                                                {
                                                    using (AttributeReference attRef = new AttributeReference())
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = unit.switchboardName;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "P_UZ":
                                                {
                                                    using (AttributeReference attRef = new AttributeReference())
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = unit.calcPower;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                        }
                                        #endregion
                                    }
                                }
                                lowestPoint = getLowestPoint(br);
                            }
                        }
                        else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
                    }
                    int nominalElectricity = 0;
                    string str = "";
                    if (unit.switchboardNominalElec.Contains(" "))
                        str = unit.switchboardNominalElec.Split(' ')[0];
                    else str = unit.switchboardNominalElec;
                    int.TryParse(str, out nominalElectricity);
                    if (nominalElectricity < 78 && unit.switchboardNominalElec != "-" && nominalElectricity != 0)
                    {
                        ids = new ObjectIdCollection();
                        filename = @"Data\SinFilter.dwg";
                        blockName = "SinFilter";
                        using (Database sourceDb = new Database(false, true))
                        {
                            if (System.IO.File.Exists(filename))
                            {
                                sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                                using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                                {
                                    BlockTable bt = (BlockTable)trans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                                    if (bt.Has(blockName))
                                        ids.Add(bt[blockName]);
                                    trans.Commit();
                                }
                            }
                            else editor.WriteMessage("Не найден файл {0}", filename);
                            if (ids.Count > 0)
                            {
                                acTrans.TransactionManager.QueueForGraphicsFlush();
                                IdMapping iMap = new IdMapping();
                                acdb.WblockCloneObjects(ids, acdb.CurrentSpaceId, iMap, DuplicateRecordCloning.Replace, false);
                                BlockTable bt = (BlockTable)acTrans.GetObject(acdb.BlockTableId, OpenMode.ForRead);
                                if (bt.Has(blockName))
                                {
                                    BlockReference br = new BlockReference(lowestPoint, bt[blockName]);
                                    br.Layer = "0";
                                    modSpace.AppendEntity(br);
                                    acTrans.AddNewlyCreatedDBObject(br, true);
                                    BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                                    foreach (ObjectId id in btr)
                                    {
                                        DBObject obj = id.GetObject(OpenMode.ForWrite);
                                        AttributeDefinition attDef = obj as AttributeDefinition;
                                        if ((attDef != null) && (!attDef.Constant) && attDef.Visible == true)
                                        {
                                            #region attributes
                                            switch (attDef.Tag)
                                            {
                                                case "NL":
                                                    {
                                                        using (AttributeReference attRef = new AttributeReference())
                                                        {
                                                            string sfNAme = currentSection + "Z" + sinFilter;
                                                            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                            attRef.TextString = sfNAme;
                                                            br.AttributeCollection.AppendAttribute(attRef);
                                                            acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                        }
                                                        break;
                                                    }
                                            }
                                            #endregion
                                        }
                                    }
                                    lowestPoint = getLowestPoint(br);
                                }
                                sinFilter++;
                            }
                            else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);

                        }
                    }
                }

                double endPointY = 125 - (currentPoint.Y - lowestPoint.Y);
                Line cableLine = new Line();
                cableLine.SetDatabaseDefaults();
                cableLine.Color = Color.FromRgb(255, 0, 0);
                cableLine.StartPoint = lowestPoint;
                cableLine.EndPoint = cableLine.StartPoint.Add(new Vector3d(0, -endPointY, 0));
                modSpace.AppendEntity(cableLine);
                acTrans.AddNewlyCreatedDBObject(cableLine, true);
                currentPoint = currentPoint.Add(new Vector3d(maxPoint.X - minPoint.X, 0, 0));

                if (!unit.consumerName.Contains("Ввод"))
                {
                    ids = new ObjectIdCollection();
                    filename = @"Data\Consumer.dwg";
                    blockName = "Consumer";
                    using (Database sourceDb = new Database(false, true))
                    {
                        if (System.IO.File.Exists(filename))
                        {
                            sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                            using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                            {
                                BlockTable bt = (BlockTable)trans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                                if (bt.Has(blockName))
                                    ids.Add(bt[blockName]);
                                trans.Commit();
                            }
                        }
                        else editor.WriteMessage("Не найден файл {0}", filename);
                        if (ids.Count > 0)
                        {
                            acTrans.TransactionManager.QueueForGraphicsFlush();
                            IdMapping iMap = new IdMapping();
                            acdb.WblockCloneObjects(ids, acdb.CurrentSpaceId, iMap, DuplicateRecordCloning.Replace, false);
                            BlockTable bt = (BlockTable)acTrans.GetObject(acdb.BlockTableId, OpenMode.ForRead);
                            if (bt.Has(blockName))
                            {
                                BlockReference br = new BlockReference(cableLine.EndPoint, bt[blockName]);
                                br.Layer = "0";
                                modSpace.AppendEntity(br);
                                acTrans.AddNewlyCreatedDBObject(br, true);
                                if (br.IsDynamicBlock)
                                {
                                    DynamicBlockReferencePropertyCollection props =
                                        br.DynamicBlockReferencePropertyCollection;

                                    foreach (DynamicBlockReferenceProperty prop in props)
                                    {
                                        object[] values = prop.GetAllowedValues();
                                        if (prop.PropertyName == "Видимость1" && !prop.ReadOnly)
                                        {
                                            if (unit.planName.Contains("AH") || unit.planName.Contains("BL") || unit.planName.Contains("FNM") || unit.planName.Contains("P"))
                                                prop.Value = values[0];
                                            else if (unit.planName.Contains("EK"))
                                                prop.Value = values[1];
                                            else if (unit.planName.Contains("FH") || unit.planName.Contains("WW") || unit.planName.Contains("FF") || unit.planName.Contains("FK") || unit.planName.Contains("FJ") || unit.planName.Contains("HPL") || unit.planName.Contains("VA"))
                                                prop.Value = values[2];
                                            else if (unit.planName.Contains("EL"))
                                                prop.Value = values[1];
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
                    }
                }
            }
        }

        static private void getSize(BlockReference bR, ref Point3d min, ref Point3d max)
        {
            DBObjectCollection objects = new DBObjectCollection();
            bR.Explode(objects);
            Extents3d ext = new Extents3d();
            for (int i = 0; i < objects.Count; i++)
            {
                if (objects[i].Bounds != null)
                {
                    Entity ent = objects[i] as Entity;
                    if (ent.Visible != false)
                        ext.AddExtents(objects[i].Bounds.Value);
                }
                else
                {
                    DBObject obj = objects[i];
                    AttributeDefinition attDef = obj as AttributeDefinition;
                    if ((attDef != null) && (!attDef.Constant) && attDef.Visible == true)
                    {
                        Point3d point = attDef.Position;
                        if (attDef.Justify == AttachmentPoint.BaseRight) point = point.Add(new Vector3d(-2 * attDef.TextString.Length, 0, 0));
                        else if (attDef.Justify == AttachmentPoint.BaseLeft) point = point.Add(new Vector3d(2 * attDef.TextString.Length, 0, 0));
                        ext.AddPoint(point);
                    }
                }
            }
            objects.Dispose();
            min = ext.MinPoint;
            max = ext.MaxPoint;
        }

        static private  Point3d getLowestPoint(BlockReference bR)
        {
            DBObjectCollection objects = new DBObjectCollection();
            bR.Explode(objects);
            Point3d point = new Point3d();
            bool first = true;
            for (int i = 0; i < objects.Count; i++)
            {
                if (objects[i].Bounds != null)
                {
                    Entity ent = objects[i] as Entity;
                    if (ent.Visible != false)
                    {
                        if (first)
                        {
                            first = false;
                            point = objects[i].Bounds.Value.MinPoint;
                        }
                        else
                        {
                            if (objects[i].Bounds.Value.MinPoint.Y < point.Y)
                                point = objects[i].Bounds.Value.MinPoint;
                        }
                    }
                }
            }
            objects.Dispose();
            return point;
        }
    }
}
