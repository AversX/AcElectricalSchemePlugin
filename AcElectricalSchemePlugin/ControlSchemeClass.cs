using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;

namespace AcElectricalSchemePlugin
{
    static class ControlSchemeClass
    {
        private static Document acDoc;
        private static Editor editor;
        private static TextStyleTable tst;
        private static int etCount = 0;
        private static int aiCount = 0;
        private static int aoCount = 0;
        private static int diCount = 0;
        private static double do24vCount = 0;
        private static int doCount = 0;
        private static int currentSheet = 8;
        private static Point3d currentPoint;
        private static List<Point3d> ets;
        private static Point3d block246;
        private static Point3d block241;
        private static Point3d block242;
        private static int currentD = 4;
        private static int[] currentPinAIAO = { 2, 2, 2, 2 };

        static public void drawControlScheme()
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            PromptResult MC = editor.GetString("\nГлавное управление?(y/n) ");
            if (MC.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else currentSheet = MC.StringResult != null && (MC.StringResult.ToUpper().Contains("Д") || MC.StringResult.ToUpper().Contains("Y")) ? currentSheet : currentSheet-1;

            PromptIntegerResult ET = editor.GetInteger("\nВведите количество ET: ");
            if (ET.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else etCount = ET.Value;

            ets = new List<Point3d>();
            for (int i = 0; i < etCount; i++)
            {
                PromptPointResult point = editor.GetPoint("\nВыберите правый верхний угол последнего модуля в ET №: " + (i + 1));
                if (point.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("Неверный ввод...");
                    return;
                }
                else ets.Add(point.Value);
            }

            PromptIntegerResult AI = editor.GetInteger("\nВведите количество модулей AI: ");
            if (AI.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else aiCount = AI.Value;

            PromptIntegerResult AO = editor.GetInteger("\nВведите количество модулей AO: ");
            if (AO.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else aoCount = AO.Value;

            if (aiCount+aoCount>8) editor.WriteMessage("\nВНИМАЕНИЕ!!! Общее количество модулей AI и AO превышает 8, будут расчитаны только первые 8 модулей.");

            PromptIntegerResult DI = editor.GetInteger("\nВведите количество модулей DI: ");
            if (DI.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else diCount = DI.Value;

            PromptResult DO24v = editor.GetString("\nВведите количество модулей DO 24V: ");
            if (DO24v.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else
            {
                double value = 0;
                double.TryParse(DO24v.StringResult, out value);
                do24vCount = value;
            }

            while (doCount < do24vCount)
            {
                if (doCount < do24vCount) editor.WriteMessage("\nКоличество модулей DO24V не должно превышать общее количество модулей DO");
                PromptIntegerResult DO = editor.GetInteger("Введите общее количество модулей DO: ");
                if (DO.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("Неверный ввод...");
                    return;
                }
                else doCount = DO.Value;
            }

            if (etCount == 2 && diCount + doCount > 10) editor.WriteMessage("\nВНИМАЕНИЕ!!! Общее количество модулей DI и DO превышает 10, будут расчитаны только первые 10 модулей.");
            else if (etCount == 3 && diCount + doCount > 16) editor.WriteMessage("\nВНИМАЕНИЕ!!! Общее количество модулей DI и DO превышает 16, будут расчитаны только первые 16 модулей.");
            
            PromptPointResult startPoint = editor.GetPoint("Выберите точку старта отрисовки");
            if (startPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else currentPoint = startPoint.Value.Add(new Vector3d(50,0,0));

            PromptPointResult aiaoPoint = editor.GetPoint("Выберите левый верхний угол блока 24.6");
            if (aiaoPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else block246 = aiaoPoint.Value.Add(new Vector3d(0, -25.9834, 0));

            PromptPointResult diDownPoint = editor.GetPoint("Выберите левый нижний угол блока 24.1");
            if (diDownPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else block241 = diDownPoint.Value.Add(new Vector3d(5.857, 0, 0));

            PromptPointResult diUpPoint = editor.GetPoint("Выберите левый нижний угол блока 24.2");
            if (diUpPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else block242 = diUpPoint.Value.Add(new Vector3d(5.857, 0, 0));

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForRead);

                    for (int i = 0; i < aiCount; i++)
                    {
                        if (i < 8)
                        {
                            insertAI(acTrans, acModSpace, acDb, currentPoint, i);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETA(acTrans, acModSpace, acDb, ets[0], i, "AI");
                            ets[0] = ets[0].Add(new Vector3d(24, 0, 0));
                        }
                    }
                    currentPoint = currentPoint.Add(new Vector3d(50, 0, 0));
                    for (int i = aiCount; i < aiCount + aoCount; i++)
                    {
                        if (i < 8 - aiCount)
                        {
                            insertAO(acTrans, acModSpace, acDb, currentPoint, i);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETA(acTrans, acModSpace, acDb, ets[0], i, "AO");
                            ets[0] = ets[0].Add(new Vector3d(24, 0, 0));
                        }
                    }
                    currentPoint = currentPoint.Add(new Vector3d(50, 0, 0));
                    insertDI1(acTrans, acModSpace, acDb, 0);
                    currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                    insertETD(acTrans, acModSpace, acDb, ets[1], 0, "DI");
                    ets[1] = ets[1].Add(new Vector3d(24, 0, 0));
                    int maxCount = 8;
                    if (etCount == 2) maxCount = 10;
                    int currentET = 1;
                    for (int i = 1; i < diCount; i++)
                    {
                        if (i < maxCount)
                        {
                            insertDI(acTrans, acModSpace, acDb, i);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DI");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                        }
                        else
                        {
                            if (i == maxCount) { currentD++; maxCount = 16; currentET++; }
                            insertDI(acTrans, acModSpace, acDb, i);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DI");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                        } 
                    }
                    //currentPoint = currentPoint.Add(new Vector3d(50, 0, 0));
                    //int currentModNum = 4 + diCount;
                    //for (double i = diCount; i < do24vCount+diCount; i+=0.5)
                    //{
                    //    if (i < maxCount)
                    //    {
                    //        insertDO24V(acTrans, acModSpace, acDb, currentModNum);
                    //        if (i % 10 == 0)
                    //        {
                    //            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModNum, "DO");
                    //            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                    //            currentModNum++;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (i == maxCount) { currentD++; maxCount = 16; currentET++; }
                    //        insertDO24V(acTrans, acModSpace, acDb, currentModNum);
                    //        if (i % 10 == 0)
                    //        {
                    //            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModNum, "DO");
                    //            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                    //            currentModNum++;
                    //        }
                    //    } 
                    //}
                    //if (do24vCount % 10 != 0)
                    //{
                    //    insertDO24V2(acTrans, acModSpace, acDb, currentModNum);
                    //    insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModNum, "DO");
                    //    ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                    //    currentModNum++;
                    //    do24vCount += 0.5;
                    //}
                    //for (int i = diCount + (int)do24vCount; i < doCount + do24vCount+ diCount; i++)
                    //{
                    //    if (i < maxCount)
                    //    {
                    //        insertDO(acTrans, acModSpace, acDb, i);
                    //        currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                    //        insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                    //        ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                    //    }
                    //    else
                    //    {
                    //        if (i == maxCount) { currentD++; maxCount = 16; currentET++; }
                    //        insertDO(acTrans, acModSpace, acDb, i);
                    //        currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                    //        insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                    //        ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                    //    } 
                    //}
                    acTrans.Commit();
                }
            }
            aiCount = 0;
            aoCount = 0;
            diCount = 0;
            doCount = 0;
            currentSheet = 8;
            currentPinAIAO = new int[] { 2, 2, 2, 2 };
            etCount = 0;
            do24vCount = 0;
            currentD = 4;
        }

        private static void insertAO(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint, int moduleNumber)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\AO.dwg";
            string blockName = "AO";
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
                        BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                        br.Layer = "0";
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);

                        Line cableLinePlus = new Line();
                        cableLinePlus.SetDatabaseDefaults();
                        cableLinePlus.Color = Color.FromRgb(255, 0, 0);
                        cableLinePlus.StartPoint = block246.Add(new Vector3d(2.4047, -(moduleNumber * 3.4408), 0));
                        cableLinePlus.EndPoint = cableLinePlus.StartPoint.Add(new Vector3d(-14, 0, 0));
                        modSpace.AppendEntity(cableLinePlus);
                        acTrans.AddNewlyCreatedDBObject(cableLinePlus, true);

                        DBText cableMarkPlus = new DBText();
                        cableMarkPlus.SetDatabaseDefaults();
                        cableMarkPlus.TextStyleId = tst["GOSTA-2.5-1"];
                        cableMarkPlus.Position = cableLinePlus.StartPoint.Add(new Vector3d(-1, 0, 0));
                        cableMarkPlus.Layer = "КИА_МАРКИРОВКА";
                        cableMarkPlus.Height = 2.5;
                        cableMarkPlus.WidthFactor = 0.8;
                        cableMarkPlus.Justify = AttachmentPoint.BottomRight;
                        cableMarkPlus.AlignmentPoint = cableMarkPlus.Position;
                        modSpace.AppendEntity(cableMarkPlus);
                        acTrans.AddNewlyCreatedDBObject(cableMarkPlus, true);

                        DBText linkPlus = new DBText();
                        linkPlus.SetDatabaseDefaults();
                        linkPlus.TextStyleId = tst["GOSTA-2.5-1"];
                        linkPlus.Position = cableLinePlus.EndPoint.Add(new Vector3d(-1, 0, 0));
                        linkPlus.Height = 2.5;
                        linkPlus.WidthFactor = 0.7;
                        linkPlus.Justify = AttachmentPoint.MiddleRight;
                        linkPlus.AlignmentPoint = linkPlus.Position;
                        linkPlus.Layer = "0";
                        modSpace.AppendEntity(linkPlus);
                        acTrans.AddNewlyCreatedDBObject(linkPlus, true);

                        Line cableLineMinus = new Line();
                        cableLineMinus.SetDatabaseDefaults();
                        cableLineMinus.Color = Color.FromRgb(0, 255, 255);
                        cableLineMinus.StartPoint = block246.Add(new Vector3d(53.4144, -(moduleNumber * 3.4408), 0));
                        cableLineMinus.EndPoint = cableLineMinus.StartPoint.Add(new Vector3d(-14, 0, 0));
                        modSpace.AppendEntity(cableLineMinus);
                        acTrans.AddNewlyCreatedDBObject(cableLineMinus, true);

                        DBText cableMarkMinus = new DBText();
                        cableMarkMinus.SetDatabaseDefaults();
                        cableMarkMinus.TextStyleId = tst["GOSTA-2.5-1"];
                        cableMarkMinus.Position = cableLineMinus.StartPoint.Add(new Vector3d(-1, 0, 0));
                        cableMarkMinus.Layer = "КИА_МАРКИРОВКА";
                        cableMarkMinus.Height = 2.5;
                        cableMarkMinus.WidthFactor = 0.8;
                        cableMarkMinus.Justify = AttachmentPoint.BottomRight;
                        cableMarkMinus.AlignmentPoint = cableMarkMinus.Position;
                        modSpace.AppendEntity(cableMarkMinus);
                        acTrans.AddNewlyCreatedDBObject(cableMarkMinus, true);

                        DBText linkMinus = new DBText();
                        linkMinus.SetDatabaseDefaults();
                        linkMinus.TextStyleId = tst["GOSTA-2.5-1"];
                        linkMinus.Position = cableLineMinus.EndPoint.Add(new Vector3d(-1, 0, 0));
                        linkMinus.Height = 2.5;
                        linkMinus.WidthFactor = 0.7;
                        linkMinus.Justify = AttachmentPoint.MiddleRight;
                        linkMinus.AlignmentPoint = linkMinus.Position;
                        linkMinus.Layer = "0";
                        modSpace.AppendEntity(linkMinus);
                        acTrans.AddNewlyCreatedDBObject(linkMinus, true);

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "3A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "+CABLE":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3L6+/1." + currentPinAIAO[0];
                                                    currentPinAIAO[0]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3L6+/2." + currentPinAIAO[1];
                                                    currentPinAIAO[1]++;
                                                }
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlus.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "-CABLE":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3M6+/3." + currentPinAIAO[2];
                                                    currentPinAIAO[2]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3M6+/4." + currentPinAIAO[3];
                                                    currentPinAIAO[3]++;
                                                }
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinus.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "LINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.5)", moduleNumber % 2 == 0 ? "1" : "2");
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPlus.TextString = string.Format("(-3.{0}:1/3.{1})", (moduleNumber + 6), currentSheet);
                                            }
                                            break;
                                        }
                                    case "LINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.5)", moduleNumber % 2 == 0 ? "3" : "4");
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinus.TextString = string.Format("(-3.{0}:20/3.{1})", (moduleNumber + 6), currentSheet);
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "WA3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        //br.ExplodeToOwnerSpace();
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }
        private static void insertAI(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint, int moduleNumber)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\AI.dwg";
            string blockName = "AI";
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
                        BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                        br.Layer = "0";
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);

                        Line cableLinePlus = new Line();
                        cableLinePlus.SetDatabaseDefaults();
                        cableLinePlus.Color = Color.FromRgb(255, 0, 0);
                        cableLinePlus.StartPoint = block246.Add(new Vector3d(2.4047, -(moduleNumber * 3.4408), 0));
                        cableLinePlus.EndPoint = cableLinePlus.StartPoint.Add(new Vector3d(-14, 0, 0));
                        modSpace.AppendEntity(cableLinePlus);
                        acTrans.AddNewlyCreatedDBObject(cableLinePlus, true);

                        DBText cableMarkPlus = new DBText();
                        cableMarkPlus.SetDatabaseDefaults();
                        cableMarkPlus.TextStyleId = tst["GOSTA-2.5-1"];
                        cableMarkPlus.Position = cableLinePlus.StartPoint.Add(new Vector3d(-1, 0, 0));
                        cableMarkPlus.Layer = "КИА_МАРКИРОВКА";
                        cableMarkPlus.Height = 2.5;
                        cableMarkPlus.WidthFactor = 0.8;
                        cableMarkPlus.Justify = AttachmentPoint.BottomRight;
                        cableMarkPlus.AlignmentPoint = cableMarkPlus.Position;
                        modSpace.AppendEntity(cableMarkPlus);
                        acTrans.AddNewlyCreatedDBObject(cableMarkPlus, true);

                        DBText linkPlus = new DBText();
                        linkPlus.SetDatabaseDefaults();
                        linkPlus.TextStyleId = tst["GOSTA-2.5-1"];
                        linkPlus.Position = cableLinePlus.EndPoint.Add(new Vector3d(-1, 0, 0));
                        linkPlus.Height = 2.5;
                        linkPlus.WidthFactor = 0.7;
                        linkPlus.Layer = "0";
                        linkPlus.Justify = AttachmentPoint.MiddleRight;
                        linkPlus.AlignmentPoint = linkPlus.Position;
                        modSpace.AppendEntity(linkPlus);
                        acTrans.AddNewlyCreatedDBObject(linkPlus, true);

                        Line cableLineMinus = new Line();
                        cableLineMinus.SetDatabaseDefaults();
                        cableLineMinus.Color = Color.FromRgb(0, 255, 255);
                        cableLineMinus.StartPoint = block246.Add(new Vector3d(53.4144, -(moduleNumber * 3.4408), 0));
                        cableLineMinus.EndPoint = cableLineMinus.StartPoint.Add(new Vector3d(-14, 0, 0));
                        modSpace.AppendEntity(cableLineMinus);
                        acTrans.AddNewlyCreatedDBObject(cableLineMinus, true);

                        DBText cableMarkMinus = new DBText();
                        cableMarkMinus.SetDatabaseDefaults();
                        cableMarkMinus.TextStyleId = tst["GOSTA-2.5-1"];
                        cableMarkMinus.Position = cableLineMinus.StartPoint.Add(new Vector3d(-1, 0, 0));
                        cableMarkMinus.Layer = "КИА_МАРКИРОВКА";
                        cableMarkMinus.Height = 2.5;
                        cableMarkMinus.WidthFactor = 0.8;
                        cableMarkMinus.Justify = AttachmentPoint.BottomRight;
                        cableMarkMinus.AlignmentPoint = cableMarkMinus.Position;
                        modSpace.AppendEntity(cableMarkMinus);
                        acTrans.AddNewlyCreatedDBObject(cableMarkMinus, true);
                       
                        DBText linkMinus = new DBText();
                        linkMinus.SetDatabaseDefaults();
                        linkMinus.TextStyleId = tst["GOSTA-2.5-1"];
                        linkMinus.Position = cableLineMinus.EndPoint.Add(new Vector3d(-1, 0, 0));
                        linkMinus.Height = 2.5;
                        linkMinus.WidthFactor = 0.7;
                        linkMinus.Justify = AttachmentPoint.MiddleRight;
                        linkMinus.AlignmentPoint = linkMinus.Position;
                        linkMinus.Layer = "0";
                        modSpace.AppendEntity(linkMinus);
                        acTrans.AddNewlyCreatedDBObject(linkMinus, true);

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "3A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "+CABLE":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3L6+/1." + currentPinAIAO[0];
                                                    currentPinAIAO[0]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3L6+/2." + currentPinAIAO[1];
                                                    currentPinAIAO[1]++;
                                                }
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlus.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "-CABLE":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3M6+/3." + currentPinAIAO[2];
                                                    currentPinAIAO[2]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3M6+/4." + currentPinAIAO[3];
                                                    currentPinAIAO[3]++;
                                                }
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinus.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "LINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.5)", moduleNumber % 2 == 0 ? "1" : "2");
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPlus.TextString = string.Format("(-3.{0}:1/3.{1})", (moduleNumber+6), currentSheet); 
                                            }
                                            break;
                                        }
                                    case "LINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.5)", moduleNumber % 2 == 0 ? "3" : "4");
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinus.TextString = string.Format("(-3.{0}:20/3.{1})", (moduleNumber + 6), currentSheet);
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "WA3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                currentSheet++;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        //br.ExplodeToOwnerSpace();
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }
        private static void insertETA(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint, int moduleNumber, string mod)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = string.Format(@"Data\ET{0}.dwg", mod);
            string blockName = string.Format("ET{0}", mod);
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
                        BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                        br.Layer = "0";
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "3A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 6);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "8X":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("8x{0} 4...20 mA", mod);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        //br.ExplodeToOwnerSpace();
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDI1(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block241;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            Circle acCircDown = new Circle();
            acCircDown.SetDatabaseDefaults();
            acCircDown.Center = cableDownPlus.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircDown.Radius = 0.75;
            modSpace.AppendEntity(acCircDown);
            acTrans.AddNewlyCreatedDBObject(acCircDown, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircDown.ObjectId);

            Hatch acHatchDown = new Hatch();
            modSpace.AppendEntity(acHatchDown);
            acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
            acHatchDown.SetDatabaseDefaults();
            acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchDown.Associative = true;
            acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchDown.EvaluateHatch(true);
            acHatchDown.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHDown = new Line();
            cableBranchHDown.SetDatabaseDefaults();
            cableBranchHDown.Layer = "0";
            cableBranchHDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchHDown.StartPoint = acCircDown.Center;
            cableBranchHDown.EndPoint = cableBranchHDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHDown, true);

            Line cableBranchDown = new Line();
            cableBranchDown.SetDatabaseDefaults();
            cableBranchDown.Layer = "0";
            cableBranchDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchDown.StartPoint = cableBranchHDown.EndPoint;
            cableBranchDown.EndPoint = cableBranchDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchDown, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusDown.Rotation = 1.57;
            cableMarkPlusDown.Height = 2.5;
            cableMarkPlusDown.WidthFactor = 0.7;
            cableMarkPlusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusDown.AlignmentPoint = cableMarkPlusDown.Position;
            cableMarkPlusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusDown, true);

            DBText linkPLusDown1 = new DBText();
            linkPLusDown1.SetDatabaseDefaults();
            linkPLusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown1.Position = cableDownPlus.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown1.Rotation = 1.57;
            linkPLusDown1.Height = 2.5;
            linkPLusDown1.WidthFactor = 0.7;
            linkPLusDown1.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown1.AlignmentPoint = linkPLusDown1.Position;
            linkPLusDown1.Layer = "0";
            linkPLusDown1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown1);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown1, true);

            DBText linkPLusDown2 = new DBText();
            linkPLusDown2.SetDatabaseDefaults();
            linkPLusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown2.Position = cableBranchDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown2.Rotation = 1.57;
            linkPLusDown2.Height = 2.5;
            linkPLusDown2.WidthFactor = 0.7;
            linkPLusDown2.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown2.AlignmentPoint = linkPLusDown2.Position;
            linkPLusDown2.Layer = "0";
            linkPLusDown2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown2);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown2, true);

            block241 = block241.Add(new Vector3d(8, 0, 0));

            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block241;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown = new DBText();
            linkMinusDown.SetDatabaseDefaults();
            linkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown.Rotation = 1.57;
            linkMinusDown.Height = 2.5;
            linkMinusDown.WidthFactor = 0.7;
            linkMinusDown.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown.AlignmentPoint = linkMinusDown.Position;
            linkMinusDown.Layer = "0";
            modSpace.AppendEntity(linkMinusDown);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown, true);

            block241 = block241.Add(new Vector3d(9.1624, 0, 0));

            Line cableBranchHPlusUp = new Line();
            cableBranchHPlusUp.SetDatabaseDefaults();
            cableBranchHPlusUp.Layer = "0";
            cableBranchHPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchHPlusUp.StartPoint = block242.Add(new Vector3d(0, -12.678, 0));
            cableBranchHPlusUp.EndPoint = cableBranchHPlusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHPlusUp, true);

            Line cableBranchPlusUp = new Line();
            cableBranchPlusUp.SetDatabaseDefaults();
            cableBranchPlusUp.Layer = "0";
            cableBranchPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchPlusUp.StartPoint = cableBranchHPlusUp.EndPoint;
            cableBranchPlusUp.EndPoint = cableBranchPlusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchPlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = block242.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusUp.Rotation = 1.57;
            cableMarkPlusUp.Height = 2.5;
            cableMarkPlusUp.WidthFactor = 0.7;
            cableMarkPlusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp.AlignmentPoint = cableMarkPlusUp.Position;
            cableMarkPlusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cableBranchPlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block242 = block242.Add(new Vector3d(8, 0, 0));

            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block242;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            Circle acCircMinusUp = new Circle();
            acCircMinusUp.SetDatabaseDefaults();
            acCircMinusUp.Center = cableMinusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircMinusUp.Radius = 0.75;
            modSpace.AppendEntity(acCircMinusUp);
            acTrans.AddNewlyCreatedDBObject(acCircMinusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircMinusUp.ObjectId);

            Hatch acHatchMinusUp = new Hatch();
            modSpace.AppendEntity(acHatchMinusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchMinusUp, true);
            acHatchMinusUp.SetDatabaseDefaults();
            acHatchMinusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchMinusUp.Associative = true;
            acHatchMinusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchMinusUp.EvaluateHatch(true);
            acHatchMinusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusUp = new Line();
            cableBranchHMinusUp.SetDatabaseDefaults();
            cableBranchHMinusUp.Layer = "0";
            cableBranchHMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusUp.StartPoint = acCircMinusUp.Center;
            cableBranchHMinusUp.EndPoint = cableBranchHMinusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusUp, true);

            Line cableBranchMinusUp = new Line();
            cableBranchMinusUp.SetDatabaseDefaults();
            cableBranchMinusUp.Layer = "0";
            cableBranchMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusUp.StartPoint = cableBranchHMinusUp.EndPoint;
            cableBranchMinusUp.EndPoint = cableBranchMinusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusUp.Rotation = 1.57;
            cableMarkMinusUp.Height = 2.5;
            cableMarkMinusUp.WidthFactor = 0.7;
            cableMarkMinusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp.AlignmentPoint = cableMarkMinusUp.Position;
            cableMarkMinusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableBranchMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);
           
            block242 = block242.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DI1.dwg";
            string blockName = "DI1";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() +"." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L"+(moduleNumber+1)+"+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":20/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G"+ currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DI2.dwg";
            blockName = "DI2";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertDI(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block241;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            Circle acCircDown = new Circle();
            acCircDown.SetDatabaseDefaults();
            acCircDown.Center = cableDownPlus.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircDown.Radius = 0.75;
            acCircDown.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(acCircDown);
            acTrans.AddNewlyCreatedDBObject(acCircDown, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircDown.ObjectId);

            Hatch acHatchDown = new Hatch();
            modSpace.AppendEntity(acHatchDown);
            acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
            acHatchDown.SetDatabaseDefaults();
            acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchDown.Associative = true;
            acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchDown.EvaluateHatch(true);
            acHatchDown.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHDown = new Line();
            cableBranchHDown.SetDatabaseDefaults();
            cableBranchHDown.Layer = "0";
            cableBranchHDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchHDown.StartPoint = acCircDown.Center;
            cableBranchHDown.EndPoint = cableBranchHDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHDown, true);

            Line cableBranchDown = new Line();
            cableBranchDown.SetDatabaseDefaults();
            cableBranchDown.Layer = "0";
            cableBranchDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchDown.StartPoint = cableBranchHDown.EndPoint;
            cableBranchDown.EndPoint = cableBranchDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchDown, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusDown.Rotation = 1.57;
            cableMarkPlusDown.Height = 2.5;
            cableMarkPlusDown.WidthFactor = 0.7;
            cableMarkPlusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusDown.AlignmentPoint = cableMarkPlusDown.Position;
            cableMarkPlusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusDown, true);

            DBText linkPLusDown1 = new DBText();
            linkPLusDown1.SetDatabaseDefaults();
            linkPLusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown1.Position = cableDownPlus.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown1.Rotation = 1.57;
            linkPLusDown1.Height = 2.5;
            linkPLusDown1.WidthFactor = 0.7;
            linkPLusDown1.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown1.AlignmentPoint = linkPLusDown1.Position;
            linkPLusDown1.Layer = "0";
            linkPLusDown1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown1);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown1, true);

            DBText linkPLusDown2 = new DBText();
            linkPLusDown2.SetDatabaseDefaults();
            linkPLusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown2.Position = cableBranchDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown2.Rotation = 1.57;
            linkPLusDown2.Height = 2.5;
            linkPLusDown2.WidthFactor = 0.7;
            linkPLusDown2.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown2.AlignmentPoint = linkPLusDown2.Position;
            linkPLusDown2.Layer = "0";
            linkPLusDown2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown2);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown2, true);

            block241 = block241.Add(new Vector3d(8, 0, 0));

            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block241;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown = new DBText();
            linkMinusDown.SetDatabaseDefaults();
            linkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown.Rotation = 1.57;
            linkMinusDown.Height = 2.5;
            linkMinusDown.WidthFactor = 0.7;
            linkMinusDown.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown.AlignmentPoint = linkMinusDown.Position;
            linkMinusDown.Layer = "0";
            modSpace.AppendEntity(linkMinusDown);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown, true);

            block241 = block241.Add(new Vector3d(9.1624, 0, 0));

            Line cablePlusUp = new Line();
            cablePlusUp.SetDatabaseDefaults();
            cablePlusUp.Layer = "0";
            cablePlusUp.Color = Color.FromRgb(255, 0, 0);
            cablePlusUp.StartPoint = block242;
            cablePlusUp.EndPoint = cablePlusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cablePlusUp);
            acTrans.AddNewlyCreatedDBObject(cablePlusUp, true);

            Circle acCircPlusUp = new Circle();
            acCircPlusUp.SetDatabaseDefaults();
            acCircPlusUp.Center = cablePlusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircPlusUp.Radius = 0.75;
            acCircPlusUp.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(acCircPlusUp);
            acTrans.AddNewlyCreatedDBObject(acCircPlusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircPlusUp.ObjectId);

            Hatch acHatchPlusUp = new Hatch();
            modSpace.AppendEntity(acHatchPlusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchPlusUp, true);
            acHatchPlusUp.SetDatabaseDefaults();
            acHatchPlusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchPlusUp.Associative = true;
            acHatchPlusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchPlusUp.EvaluateHatch(true);
            acHatchPlusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHPlusUp = new Line();
            cableBranchHPlusUp.SetDatabaseDefaults();
            cableBranchHPlusUp.Layer = "0";
            cableBranchHPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchHPlusUp.StartPoint = acCircPlusUp.Center;
            cableBranchHPlusUp.EndPoint = cableBranchHPlusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHPlusUp, true);

            Line cableBranchPlusUp = new Line();
            cableBranchPlusUp.SetDatabaseDefaults();
            cableBranchPlusUp.Layer = "0";
            cableBranchPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchPlusUp.StartPoint = cableBranchHPlusUp.EndPoint;
            cableBranchPlusUp.EndPoint = cableBranchPlusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchPlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = cablePlusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusUp.Rotation = 1.57;
            cableMarkPlusUp.Height = 2.5;
            cableMarkPlusUp.WidthFactor = 0.7;
            cableMarkPlusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp.AlignmentPoint = cableMarkPlusUp.Position;
            cableMarkPlusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp, true);

            DBText linkPLusUp1 = new DBText();
            linkPLusUp1.SetDatabaseDefaults();
            linkPLusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp1.Position = cablePlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp1.Rotation = 1.57;
            linkPLusUp1.Height = 2.5;
            linkPLusUp1.WidthFactor = 0.7;
            linkPLusUp1.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp1.AlignmentPoint = linkPLusUp1.Position;
            linkPLusUp1.Layer = "0";
            linkPLusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp1);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp1, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cableBranchPlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block242 = block242.Add(new Vector3d(8, 0, 0));

            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block242;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            Circle acCircMinusUp = new Circle();
            acCircMinusUp.SetDatabaseDefaults();
            acCircMinusUp.Center = cableMinusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircMinusUp.Radius = 0.75;
            acCircMinusUp.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(acCircMinusUp);
            acTrans.AddNewlyCreatedDBObject(acCircMinusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircMinusUp.ObjectId);

            Hatch acHatchMinusUp = new Hatch();
            modSpace.AppendEntity(acHatchMinusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchMinusUp, true);
            acHatchMinusUp.SetDatabaseDefaults();
            acHatchMinusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchMinusUp.Associative = true;
            acHatchMinusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchMinusUp.EvaluateHatch(true);
            acHatchMinusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusUp = new Line();
            cableBranchHMinusUp.SetDatabaseDefaults();
            cableBranchHMinusUp.Layer = "0";
            cableBranchHMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusUp.StartPoint = acCircMinusUp.Center;
            cableBranchHMinusUp.EndPoint = cableBranchHMinusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusUp, true);

            Line cableBranchMinusUp = new Line();
            cableBranchMinusUp.SetDatabaseDefaults();
            cableBranchMinusUp.Layer = "0";
            cableBranchMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusUp.StartPoint = cableBranchHMinusUp.EndPoint;
            cableBranchMinusUp.EndPoint = cableBranchMinusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusUp.Rotation = 1.57;
            cableMarkMinusUp.Height = 2.5;
            cableMarkMinusUp.WidthFactor = 0.7;
            cableMarkMinusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp.AlignmentPoint = cableMarkMinusUp.Position;
            cableMarkMinusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableBranchMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);

            block242 = block242.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DIF.dwg";
            string blockName = "DIF";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":20/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp1.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
            ids = new ObjectIdCollection();
            filename = @"Data\DIS.dwg";
            blockName = "DIS";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertDO24V(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block241;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            Circle acCircDown = new Circle();
            acCircDown.SetDatabaseDefaults();
            acCircDown.Center = cableDownPlus.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircDown.Radius = 0.75;
            modSpace.AppendEntity(acCircDown);
            acTrans.AddNewlyCreatedDBObject(acCircDown, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircDown.ObjectId);

            Hatch acHatchDown = new Hatch();
            modSpace.AppendEntity(acHatchDown);
            acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
            acHatchDown.SetDatabaseDefaults();
            acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchDown.Associative = true;
            acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchDown.EvaluateHatch(true);
            acHatchDown.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHDown = new Line();
            cableBranchHDown.SetDatabaseDefaults();
            cableBranchHDown.Layer = "0";
            cableBranchHDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchHDown.StartPoint = acCircDown.Center;
            cableBranchHDown.EndPoint = cableBranchHDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHDown, true);

            Line cableBranchDown = new Line();
            cableBranchDown.SetDatabaseDefaults();
            cableBranchDown.Layer = "0";
            cableBranchDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchDown.StartPoint = cableBranchHDown.EndPoint;
            cableBranchDown.EndPoint = cableBranchDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchDown, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusDown.Rotation = 1.57;
            cableMarkPlusDown.Height = 2.5;
            cableMarkPlusDown.WidthFactor = 0.7;
            cableMarkPlusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusDown.AlignmentPoint = cableMarkPlusDown.Position;
            cableMarkPlusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusDown, true);

            DBText linkPLusDown1 = new DBText();
            linkPLusDown1.SetDatabaseDefaults();
            linkPLusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown1.Position = cableDownPlus.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown1.Rotation = 1.57;
            linkPLusDown1.Height = 2.5;
            linkPLusDown1.WidthFactor = 0.7;
            linkPLusDown1.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown1.AlignmentPoint = linkPLusDown1.Position;
            linkPLusDown1.Layer = "0";
            linkPLusDown1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown1);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown1, true);

            DBText linkPLusDown2 = new DBText();
            linkPLusDown2.SetDatabaseDefaults();
            linkPLusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown2.Position = cableBranchDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown2.Rotation = 1.57;
            linkPLusDown2.Height = 2.5;
            linkPLusDown2.WidthFactor = 0.7;
            linkPLusDown2.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown2.AlignmentPoint = linkPLusDown2.Position;
            linkPLusDown2.Layer = "0";
            linkPLusDown2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown2);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown2, true);

            block241 = block241.Add(new Vector3d(8, 0, 0));

            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block241;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown = new DBText();
            linkMinusDown.SetDatabaseDefaults();
            linkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown.Rotation = 1.57;
            linkMinusDown.Height = 2.5;
            linkMinusDown.WidthFactor = 0.7;
            linkMinusDown.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown.AlignmentPoint = linkMinusDown.Position;
            linkMinusDown.Layer = "0";
            modSpace.AppendEntity(linkMinusDown);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown, true);

            block241 = block241.Add(new Vector3d(9.1624, 0, 0));

            Line cableBranchHPlusUp = new Line();
            cableBranchHPlusUp.SetDatabaseDefaults();
            cableBranchHPlusUp.Layer = "0";
            cableBranchHPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchHPlusUp.StartPoint = block242.Add(new Vector3d(0, -12.678, 0));
            cableBranchHPlusUp.EndPoint = cableBranchHPlusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHPlusUp, true);

            Line cableBranchPlusUp = new Line();
            cableBranchPlusUp.SetDatabaseDefaults();
            cableBranchPlusUp.Layer = "0";
            cableBranchPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchPlusUp.StartPoint = cableBranchHPlusUp.EndPoint;
            cableBranchPlusUp.EndPoint = cableBranchPlusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchPlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = block242.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusUp.Rotation = 1.57;
            cableMarkPlusUp.Height = 2.5;
            cableMarkPlusUp.WidthFactor = 0.7;
            cableMarkPlusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp.AlignmentPoint = cableMarkPlusUp.Position;
            cableMarkPlusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cableBranchPlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block242 = block242.Add(new Vector3d(8, 0, 0));

            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block242;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            Circle acCircMinusUp = new Circle();
            acCircMinusUp.SetDatabaseDefaults();
            acCircMinusUp.Center = cableMinusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircMinusUp.Radius = 0.75;
            modSpace.AppendEntity(acCircMinusUp);
            acTrans.AddNewlyCreatedDBObject(acCircMinusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircMinusUp.ObjectId);

            Hatch acHatchMinusUp = new Hatch();
            modSpace.AppendEntity(acHatchMinusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchMinusUp, true);
            acHatchMinusUp.SetDatabaseDefaults();
            acHatchMinusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchMinusUp.Associative = true;
            acHatchMinusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchMinusUp.EvaluateHatch(true);
            acHatchMinusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusUp = new Line();
            cableBranchHMinusUp.SetDatabaseDefaults();
            cableBranchHMinusUp.Layer = "0";
            cableBranchHMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusUp.StartPoint = acCircMinusUp.Center;
            cableBranchHMinusUp.EndPoint = cableBranchHMinusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusUp, true);

            Line cableBranchMinusUp = new Line();
            cableBranchMinusUp.SetDatabaseDefaults();
            cableBranchMinusUp.Layer = "0";
            cableBranchMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusUp.StartPoint = cableBranchHMinusUp.EndPoint;
            cableBranchMinusUp.EndPoint = cableBranchMinusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusUp.Rotation = 1.57;
            cableMarkMinusUp.Height = 2.5;
            cableMarkMinusUp.WidthFactor = 0.7;
            cableMarkMinusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp.AlignmentPoint = cableMarkMinusUp.Position;
            cableMarkMinusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableBranchMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);

            block242 = block242.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DI1.dwg";
            string blockName = "DI1";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":20/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DI2.dwg";
            blockName = "DI2";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertDO24V2(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block241;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            Circle acCircDown = new Circle();
            acCircDown.SetDatabaseDefaults();
            acCircDown.Center = cableDownPlus.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircDown.Radius = 0.75;
            modSpace.AppendEntity(acCircDown);
            acTrans.AddNewlyCreatedDBObject(acCircDown, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircDown.ObjectId);

            Hatch acHatchDown = new Hatch();
            modSpace.AppendEntity(acHatchDown);
            acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
            acHatchDown.SetDatabaseDefaults();
            acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchDown.Associative = true;
            acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchDown.EvaluateHatch(true);
            acHatchDown.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHDown = new Line();
            cableBranchHDown.SetDatabaseDefaults();
            cableBranchHDown.Layer = "0";
            cableBranchHDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchHDown.StartPoint = acCircDown.Center;
            cableBranchHDown.EndPoint = cableBranchHDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHDown, true);

            Line cableBranchDown = new Line();
            cableBranchDown.SetDatabaseDefaults();
            cableBranchDown.Layer = "0";
            cableBranchDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchDown.StartPoint = cableBranchHDown.EndPoint;
            cableBranchDown.EndPoint = cableBranchDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchDown, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusDown.Rotation = 1.57;
            cableMarkPlusDown.Height = 2.5;
            cableMarkPlusDown.WidthFactor = 0.7;
            cableMarkPlusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusDown.AlignmentPoint = cableMarkPlusDown.Position;
            cableMarkPlusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusDown, true);

            DBText linkPLusDown1 = new DBText();
            linkPLusDown1.SetDatabaseDefaults();
            linkPLusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown1.Position = cableDownPlus.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown1.Rotation = 1.57;
            linkPLusDown1.Height = 2.5;
            linkPLusDown1.WidthFactor = 0.7;
            linkPLusDown1.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown1.AlignmentPoint = linkPLusDown1.Position;
            linkPLusDown1.Layer = "0";
            linkPLusDown1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown1);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown1, true);

            DBText linkPLusDown2 = new DBText();
            linkPLusDown2.SetDatabaseDefaults();
            linkPLusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown2.Position = cableBranchDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown2.Rotation = 1.57;
            linkPLusDown2.Height = 2.5;
            linkPLusDown2.WidthFactor = 0.7;
            linkPLusDown2.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown2.AlignmentPoint = linkPLusDown2.Position;
            linkPLusDown2.Layer = "0";
            linkPLusDown2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown2);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown2, true);

            block241 = block241.Add(new Vector3d(8, 0, 0));

            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block241;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown = new DBText();
            linkMinusDown.SetDatabaseDefaults();
            linkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown.Rotation = 1.57;
            linkMinusDown.Height = 2.5;
            linkMinusDown.WidthFactor = 0.7;
            linkMinusDown.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown.AlignmentPoint = linkMinusDown.Position;
            linkMinusDown.Layer = "0";
            modSpace.AppendEntity(linkMinusDown);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown, true);

            block241 = block241.Add(new Vector3d(9.1624, 0, 0));

            Line cableBranchHPlusUp = new Line();
            cableBranchHPlusUp.SetDatabaseDefaults();
            cableBranchHPlusUp.Layer = "0";
            cableBranchHPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchHPlusUp.StartPoint = block242.Add(new Vector3d(0, -12.678, 0));
            cableBranchHPlusUp.EndPoint = cableBranchHPlusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHPlusUp, true);

            Line cableBranchPlusUp = new Line();
            cableBranchPlusUp.SetDatabaseDefaults();
            cableBranchPlusUp.Layer = "0";
            cableBranchPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchPlusUp.StartPoint = cableBranchHPlusUp.EndPoint;
            cableBranchPlusUp.EndPoint = cableBranchPlusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchPlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = block242.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusUp.Rotation = 1.57;
            cableMarkPlusUp.Height = 2.5;
            cableMarkPlusUp.WidthFactor = 0.7;
            cableMarkPlusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp.AlignmentPoint = cableMarkPlusUp.Position;
            cableMarkPlusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cableBranchPlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block242 = block242.Add(new Vector3d(8, 0, 0));

            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block242;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            Circle acCircMinusUp = new Circle();
            acCircMinusUp.SetDatabaseDefaults();
            acCircMinusUp.Center = cableMinusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircMinusUp.Radius = 0.75;
            modSpace.AppendEntity(acCircMinusUp);
            acTrans.AddNewlyCreatedDBObject(acCircMinusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircMinusUp.ObjectId);

            Hatch acHatchMinusUp = new Hatch();
            modSpace.AppendEntity(acHatchMinusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchMinusUp, true);
            acHatchMinusUp.SetDatabaseDefaults();
            acHatchMinusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchMinusUp.Associative = true;
            acHatchMinusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchMinusUp.EvaluateHatch(true);
            acHatchMinusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusUp = new Line();
            cableBranchHMinusUp.SetDatabaseDefaults();
            cableBranchHMinusUp.Layer = "0";
            cableBranchHMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusUp.StartPoint = acCircMinusUp.Center;
            cableBranchHMinusUp.EndPoint = cableBranchHMinusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusUp, true);

            Line cableBranchMinusUp = new Line();
            cableBranchMinusUp.SetDatabaseDefaults();
            cableBranchMinusUp.Layer = "0";
            cableBranchMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusUp.StartPoint = cableBranchHMinusUp.EndPoint;
            cableBranchMinusUp.EndPoint = cableBranchMinusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusUp.Rotation = 1.57;
            cableMarkMinusUp.Height = 2.5;
            cableMarkMinusUp.WidthFactor = 0.7;
            cableMarkMinusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp.AlignmentPoint = cableMarkMinusUp.Position;
            cableMarkMinusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableBranchMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);

            block242 = block242.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DI1.dwg";
            string blockName = "DI1";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":20/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DI2.dwg";
            blockName = "DI2";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertDO(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block241;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            Circle acCircDown = new Circle();
            acCircDown.SetDatabaseDefaults();
            acCircDown.Center = cableDownPlus.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircDown.Radius = 0.75;
            modSpace.AppendEntity(acCircDown);
            acTrans.AddNewlyCreatedDBObject(acCircDown, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircDown.ObjectId);

            Hatch acHatchDown = new Hatch();
            modSpace.AppendEntity(acHatchDown);
            acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
            acHatchDown.SetDatabaseDefaults();
            acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchDown.Associative = true;
            acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchDown.EvaluateHatch(true);
            acHatchDown.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHDown = new Line();
            cableBranchHDown.SetDatabaseDefaults();
            cableBranchHDown.Layer = "0";
            cableBranchHDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchHDown.StartPoint = acCircDown.Center;
            cableBranchHDown.EndPoint = cableBranchHDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHDown, true);

            Line cableBranchDown = new Line();
            cableBranchDown.SetDatabaseDefaults();
            cableBranchDown.Layer = "0";
            cableBranchDown.Color = Color.FromRgb(255, 0, 0);
            cableBranchDown.StartPoint = cableBranchHDown.EndPoint;
            cableBranchDown.EndPoint = cableBranchDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchDown, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusDown.Rotation = 1.57;
            cableMarkPlusDown.Height = 2.5;
            cableMarkPlusDown.WidthFactor = 0.7;
            cableMarkPlusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusDown.AlignmentPoint = cableMarkPlusDown.Position;
            cableMarkPlusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusDown, true);

            DBText linkPLusDown1 = new DBText();
            linkPLusDown1.SetDatabaseDefaults();
            linkPLusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown1.Position = cableDownPlus.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown1.Rotation = 1.57;
            linkPLusDown1.Height = 2.5;
            linkPLusDown1.WidthFactor = 0.7;
            linkPLusDown1.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown1.AlignmentPoint = linkPLusDown1.Position;
            linkPLusDown1.Layer = "0";
            linkPLusDown1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown1);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown1, true);

            DBText linkPLusDown2 = new DBText();
            linkPLusDown2.SetDatabaseDefaults();
            linkPLusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusDown2.Position = cableBranchDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusDown2.Rotation = 1.57;
            linkPLusDown2.Height = 2.5;
            linkPLusDown2.WidthFactor = 0.7;
            linkPLusDown2.Justify = AttachmentPoint.MiddleRight;
            linkPLusDown2.AlignmentPoint = linkPLusDown2.Position;
            linkPLusDown2.Layer = "0";
            linkPLusDown2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusDown2);
            acTrans.AddNewlyCreatedDBObject(linkPLusDown2, true);

            block241 = block241.Add(new Vector3d(8, 0, 0));

            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block241;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown = new DBText();
            linkMinusDown.SetDatabaseDefaults();
            linkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown.Rotation = 1.57;
            linkMinusDown.Height = 2.5;
            linkMinusDown.WidthFactor = 0.7;
            linkMinusDown.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown.AlignmentPoint = linkMinusDown.Position;
            linkMinusDown.Layer = "0";
            modSpace.AppendEntity(linkMinusDown);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown, true);

            block241 = block241.Add(new Vector3d(9.1624, 0, 0));

            Line cableBranchHPlusUp = new Line();
            cableBranchHPlusUp.SetDatabaseDefaults();
            cableBranchHPlusUp.Layer = "0";
            cableBranchHPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchHPlusUp.StartPoint = block242.Add(new Vector3d(0, -12.678, 0));
            cableBranchHPlusUp.EndPoint = cableBranchHPlusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHPlusUp, true);

            Line cableBranchPlusUp = new Line();
            cableBranchPlusUp.SetDatabaseDefaults();
            cableBranchPlusUp.Layer = "0";
            cableBranchPlusUp.Color = Color.FromRgb(255, 0, 0);
            cableBranchPlusUp.StartPoint = cableBranchHPlusUp.EndPoint;
            cableBranchPlusUp.EndPoint = cableBranchPlusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchPlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = block242.Add(new Vector3d(-1, -1, 0));
            cableMarkPlusUp.Rotation = 1.57;
            cableMarkPlusUp.Height = 2.5;
            cableMarkPlusUp.WidthFactor = 0.7;
            cableMarkPlusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp.AlignmentPoint = cableMarkPlusUp.Position;
            cableMarkPlusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cableBranchPlusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block242 = block242.Add(new Vector3d(8, 0, 0));

            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block242;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            Circle acCircMinusUp = new Circle();
            acCircMinusUp.SetDatabaseDefaults();
            acCircMinusUp.Center = cableMinusUp.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCircMinusUp.Radius = 0.75;
            modSpace.AppendEntity(acCircMinusUp);
            acTrans.AddNewlyCreatedDBObject(acCircMinusUp, true);

            acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCircMinusUp.ObjectId);

            Hatch acHatchMinusUp = new Hatch();
            modSpace.AppendEntity(acHatchMinusUp);
            acTrans.AddNewlyCreatedDBObject(acHatchMinusUp, true);
            acHatchMinusUp.SetDatabaseDefaults();
            acHatchMinusUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchMinusUp.Associative = true;
            acHatchMinusUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatchMinusUp.EvaluateHatch(true);
            acHatchMinusUp.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusUp = new Line();
            cableBranchHMinusUp.SetDatabaseDefaults();
            cableBranchHMinusUp.Layer = "0";
            cableBranchHMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusUp.StartPoint = acCircMinusUp.Center;
            cableBranchHMinusUp.EndPoint = cableBranchHMinusUp.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusUp, true);

            Line cableBranchMinusUp = new Line();
            cableBranchMinusUp.SetDatabaseDefaults();
            cableBranchMinusUp.Layer = "0";
            cableBranchMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusUp.StartPoint = cableBranchHMinusUp.EndPoint;
            cableBranchMinusUp.EndPoint = cableBranchMinusUp.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(-1, -1, 0));
            cableMarkMinusUp.Rotation = 1.57;
            cableMarkMinusUp.Height = 2.5;
            cableMarkMinusUp.WidthFactor = 0.7;
            cableMarkMinusUp.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp.AlignmentPoint = cableMarkMinusUp.Position;
            cableMarkMinusUp.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableBranchMinusUp.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);

            block242 = block242.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DI1.dwg";
            string blockName = "DI1";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":20/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-1G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DI2.dwg";
            blockName = "DI2";
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

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 1) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1M" + (moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-WA" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XR":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2XR" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "G":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD1":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "VD2":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertETD(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint, int moduleNumber, string mod)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\ETD.dwg";
            string blockName = "ETD";
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
                        BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                        br.Layer = "0";
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);

                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                switch (attDef.Tag)
                                {
                                    case "4A":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-4A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP4." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "32X":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("32x{0}", mod);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        //br.ExplodeToOwnerSpace();
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }
    }
}