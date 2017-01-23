using System.Collections.Generic;
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
        private static int currentSheet = 9;
        private static int sheetNum246 = 6;
        private static bool mainControl = true;
        private static bool gas = true;
        private static Point3d currentPoint;
        private static List<Point3d> ets;
        private static Point3d link1; //4.4.1-11
        private static Point3d link2; //KV2-11
        private static Point3d link3; //4.4.1-10
        private static Point3d link4; //4.4.1-12
        private static Point3d link5; //4.4.1-13
        private static Point3d link6; //4.4.1-14
        private static Point3d link7; //4.4.1-15
        private static Point3d link8; //4.4.1-16
        private static Point3d block246;
        private static Point3d block241;
        private static Point3d block242;
        private static Point3d block243;
        private static Point3d block244;
        private static int currentD = 4;
        private static int[] currentPinAIAO = { 2, 2, 2, 2 };
        private static bool firstAI = true;
        private static bool firstAO = true;
        private static bool firstDI = true;
        private static bool firstDO = true;
        private static int current244pin = 1;
        private static string title;

        private static int lastSheetDI;
        private static int lastSheetDO;
        private static int lastDNumDI;
        private static int lastDNumDO;
        private static int lastModNumDI;
        private static int lastModNumDO;
        #region gasCablesDI
        private static AttributeReference gasCableDI1P;
        private static AttributeReference gasCableDI1M;
        private static AttributeReference gasCableDI2P;
        private static AttributeReference gasCableDI2M;
        private static AttributeReference gasCableDI3P;
        private static AttributeReference gasCableDI3M;
        private static AttributeReference gasCableDI4P;
        private static AttributeReference gasCableDI4M;
        private static AttributeReference gasCableDI5P;
        private static AttributeReference gasCableDI5M;
        private static AttributeReference gasCableDI6P;
        private static AttributeReference gasCableDI6M;
        #endregion
        #region gasLinksDI
        private static AttributeReference gasLinkDI1P;
        private static AttributeReference gasLinkDI1M;
        private static AttributeReference gasLinkDI2P;
        private static AttributeReference gasLinkDI2M;
        private static AttributeReference gasLinkDI3P;
        private static AttributeReference gasLinkDI3M;
        private static AttributeReference gasLinkDI4P;
        private static AttributeReference gasLinkDI4M;
        private static AttributeReference gasLinkDI5P;
        private static AttributeReference gasLinkDI5M;
        private static AttributeReference gasLinkDI6P;
        private static AttributeReference gasLinkDI6M;
        #endregion
        #region gasCablesDO
        private static AttributeReference gasCableDO1P;
        private static AttributeReference gasCableDO1M;
        private static AttributeReference gasCableDO2P;
        private static AttributeReference gasCableDO2M;
        private static AttributeReference gasCableDO3P;
        private static AttributeReference gasCableDO3M;
        private static AttributeReference gasCableDO4P;
        private static AttributeReference gasCableDO4M;
        #endregion
        #region gasLinksDO
        private static AttributeReference gasLinkDO1P;
        private static AttributeReference gasLinkDO1M;
        private static AttributeReference gasLinkDO2P;
        private static AttributeReference gasLinkDO2M;
        private static AttributeReference gasLinkDO3P;
        private static AttributeReference gasLinkDO3M;
        private static AttributeReference gasLinkDO4P;
        private static AttributeReference gasLinkDO4M;
        #endregion

        static public void drawControlScheme()
        {
            aiCount = 0;
            aoCount = 0;
            diCount = 0;
            doCount = 0;
            currentSheet = 9;
            sheetNum246 = 6;
            currentPinAIAO = new int[] { 2, 2, 2, 2 };
            etCount = 0;
            do24vCount = 0;
            currentD = 4;
            mainControl = true;
            firstAI = true;
            firstAO = true;
            firstDI = true;
            firstDO = true;
            current244pin = 1;
            title="";

            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            PromptResult MC = editor.GetString("\nГлавное управление?(y/n) ");
            if (MC.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else
                if (MC.StringResult != null)
                    if (!(MC.StringResult.ToUpper().Contains("Д") || MC.StringResult.ToUpper().Contains("Y")))
                    {
                        mainControl = false;
                        gas = false;
                        currentSheet--;
                        sheetNum246--;
                    }

            if (mainControl)
            {
                PromptResult G = editor.GetString("\nС загазованность?(y/n) ");
                if (G.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("Неверный ввод...");
                    return;
                }
                else
                    if (G.StringResult != null)
                        if (!(G.StringResult.ToUpper().Contains("Д") || G.StringResult.ToUpper().Contains("Y")))
                            gas = false;
            }

            PromptIntegerResult ET = editor.GetInteger("\nВведите количество ET: ");
            if (ET.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else etCount = ET.Value;

            ets = new List<Point3d>();

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

            if (aiCount + aoCount > 10) editor.WriteMessage("\nВНИМАЕНИЕ!!! Общее количество модулей AI и AO превышает 10, будут расчитаны только первые 10 модулей.");

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

            PromptPointResult startPoint = editor.GetPoint("Выберите левый верхний угол первого листа в чертеже");
            if (startPoint.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else setPoints(startPoint.Value);

            currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
            block246 = block246.Add(new Vector3d(0, -25.9834, 0));
            block241 = block241.Add(new Vector3d(5.857, 0, 0));
            block242 = block242.Add(new Vector3d(5.857, 0, 0));
            block243 = block243.Add(new Vector3d(5.857, 0, 0));
            block244 = block244.Add(new Vector3d(17.1729, 0, 0));

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForRead);

                    TypedValue[] filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue(0, "INSERT");
                    filterlist[1] = new TypedValue(2, "1234");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    PromptSelectionResult selRes =  editor.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nВ болванке не найден блок \"1234\"");
                        return;
                    }
                    else
                    {
                        BlockReference oEnt = (BlockReference)acTrans.GetObject(selRes.Value[0].ObjectId, OpenMode.ForRead);
                        foreach (ObjectId id in oEnt.AttributeCollection)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForWrite);
                            AttributeReference attRef = obj as AttributeReference;
                            if ((attRef != null))
                            {
                                switch (attRef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            title = attRef.TextString;
                                            break;
                                        }
                                }
                            }
                        }
                    }

                    filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue(0, "INSERT");
                    filterlist[1] = new TypedValue(2, "frm6");
                    filter = new SelectionFilter(filterlist);
                    selRes = editor.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nВ болванке не найден блок \"1234\"");
                        return;
                    }
                    else
                    {
                        SelectionSet set = selRes.Value;
                        foreach (ObjectId Id in set.GetObjectIds())
                        {
                            BlockReference oEnt = (BlockReference)acTrans.GetObject(Id, OpenMode.ForWrite);
                            foreach (ObjectId id in oEnt.AttributeCollection)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForWrite);
                                AttributeReference attRef = obj as AttributeReference;
                                if ((attRef != null))
                                {
                                    switch (attRef.Tag)
                                    {
                                        case "TITLE":
                                            {
                                                attRef.TextString = title;
                                                break;
                                            }
                                    }
                                }
                            }
                        }
                    }

                    Table table = insertTable(acTrans, acModSpace, acDb, currentPoint);
                    addRowsToTable(acTrans, acModSpace, acDb, currentPoint, table);
                    for (int i = 0; i < aiCount; i++)
                    {
                        if (i < 10)
                        {
                            insertAI(acTrans, acModSpace, acDb, currentPoint, i);
                            addRowsToTable(acTrans, acModSpace, acDb, currentPoint, table);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETA(acTrans, acModSpace, acDb, ets[0], i, "AI");
                            ets[0] = ets[0].Add(new Vector3d(24, 0, 0));
                        }
                    }
                    currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
                    table = insertTable(acTrans, acModSpace, acDb, currentPoint);
                    for (int i = aiCount; i < aiCount + aoCount; i++)
                    {
                        if (i < 10)
                        {
                            insertAO(acTrans, acModSpace, acDb, currentPoint, i);
                            //insertAI(acTrans, acModSpace, acDb, currentPoint, i);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETA(acTrans, acModSpace, acDb, ets[0], i, "AO");
                            ets[0] = ets[0].Add(new Vector3d(24, 0, 0));
                        }
                    }
                    currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
                    insertDI1(acTrans, acModSpace, acDb, 0);
                    currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                    insertETD(acTrans, acModSpace, acDb, ets[1], 0, "DI");
                    ets[1] = ets[1].Add(new Vector3d(24, 0, 0));
                    int maxCount = 8;
                    if (etCount == 2) maxCount = 10;
                    int currentET = 1;
                    if (mainControl)
                    {
                        for (int i = 1; i < diCount - 1; i++)
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
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                }
                            }
                        }
                        if (diCount - 1 < maxCount)
                        {
                            insertDIGas(acTrans, acModSpace, acDb, diCount - 1);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], diCount - 1, "DI");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                        }
                        else
                        {
                            if (currentET != 2)
                            {
                                if (etCount == 3)
                                {
                                    currentET++;
                                    currentD++;
                                    maxCount = 16;
                                    insertDIGas(acTrans, acModSpace, acDb, diCount - 1);
                                    currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                    insertETD(acTrans, acModSpace, acDb, ets[currentET], diCount - 1, "DI");
                                    ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                                }
                                else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                            }
                        }

                        currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
                        int currentModule = diCount;
                        for (int i = diCount; i < Math.Truncate(do24vCount) + diCount; i++)
                        {
                            if (i < maxCount)
                            {
                                insertDO24V(acTrans, acModSpace, acDb, currentModule);
                                currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                                ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                                currentModule++;
                            }
                            else
                            {
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else
                                    {
                                        editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                        break;
                                    }
                                }
                            }
                        }
                        if (do24vCount - Math.Truncate(do24vCount) != 0)
                        {
                            insertDO24Vhalf(acTrans, acModSpace, acDb, currentModule);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModule, "DO");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            currentModule++;
                            do24vCount += 0.5;
                        }
                        else
                        {
                            insertDO1(acTrans, acModSpace, acDb, currentModule);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModule, "DO");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            currentModule++;
                            doCount--;
                        }
                        for (int i = currentModule; i < currentModule + doCount - Math.Truncate(do24vCount) - 1; i++)
                        {
                            if (i < maxCount)
                            {
                                insertDO(acTrans, acModSpace, acDb, i);
                                currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                                ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            }
                            else
                            {
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                }
                            }
                        }
                        int x = currentModule + doCount - (int)Math.Truncate(do24vCount);
                        if (x - 1 < maxCount)
                        {
                            insertDOGas(acTrans, acModSpace, acDb, x - 1, gas);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], x - 1, "DO");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                        }
                        else
                        {
                            if (currentET != 2)
                            {
                                if (etCount == 3)
                                {
                                    currentET++;
                                    currentD++;
                                    maxCount = 16;
                                    insertDOGas(acTrans, acModSpace, acDb, x - 1, gas);
                                    currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                    insertETD(acTrans, acModSpace, acDb, ets[currentET], x - 1, "DO");
                                    ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                                }
                                else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                            }
                        }
                    }
                    else
                    {
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
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                }
                            }
                        }
                        currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
                        int currentModule = diCount;
                        for (int i = diCount; i < Math.Truncate(do24vCount) + diCount; i++)
                        {
                            if (i < maxCount)
                            {
                                insertDO24V(acTrans, acModSpace, acDb, currentModule);
                                currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                                ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                                currentModule++;
                            }
                            else
                            {
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else
                                    {
                                        editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                        break;
                                    }
                                }
                            }
                        }
                        if (do24vCount - Math.Truncate(do24vCount) != 0)
                        {
                            insertDO24Vhalf(acTrans, acModSpace, acDb, currentModule);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModule, "DO");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            currentModule++;
                            do24vCount += 0.5;
                        }
                        else
                        {
                            insertDO1(acTrans, acModSpace, acDb, currentModule);
                            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                            insertETD(acTrans, acModSpace, acDb, ets[currentET], currentModule, "DO");
                            ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            currentModule++;
                            doCount--;
                        }
                        for (int i = currentModule; i < currentModule + doCount - Math.Truncate(do24vCount); i++)
                        {
                            if (i < maxCount)
                            {
                                insertDO(acTrans, acModSpace, acDb, i);
                                currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                                insertETD(acTrans, acModSpace, acDb, ets[currentET], i, "DO");
                                ets[currentET] = ets[currentET].Add(new Vector3d(24, 0, 0));
                            }
                            else
                            {
                                if (currentET != 2)
                                {
                                    if (etCount == 3)
                                    {
                                        currentET++;
                                        currentD++;
                                        maxCount = 16;
                                        i--;
                                    }
                                    else editor.WriteMessage("Количество модулей DI/DO превышает допустимое");
                                }
                            }
                        }
                    }
                    currentPoint = currentPoint.Add(new Vector3d(500, 0, 0));
                    insertSpec(acTrans, acModSpace, acDb, currentPoint);
                    
                    filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue(0, "INSERT");
                    filterlist[1] = new TypedValue(2, "1234");
                    filter = new SelectionFilter(filterlist);
                    selRes = editor.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nВ болванке не найден блок \"1234\"");
                        return;
                    }
                    else
                    {
                        BlockReference oEnt = (BlockReference)acTrans.GetObject(selRes.Value[0].ObjectId, OpenMode.ForRead);
                        foreach (ObjectId id in oEnt.AttributeCollection)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForWrite);
                            AttributeReference attRef = obj as AttributeReference;
                            if ((attRef != null))
                            {
                                switch (attRef.Tag)
                                {
                                    case "SHS":
                                        {
                                            attRef.TextString = (currentSheet+8).ToString();
                                            break;
                                        }
                                }
                            }
                        }
                    }

                    acTrans.Commit();
                }
            }
            aiCount = 0;
            aoCount = 0;
            diCount = 0;
            doCount = 0;
            currentSheet = 9;
            sheetNum246 = 6;
            currentPinAIAO = new int[] { 2, 2, 2, 2 };
            etCount = 0;
            do24vCount = 0;
            currentD = 4;
            mainControl = true;
            firstAI = true;
            firstAO = true;
            firstDI = true;
            firstDO = true;
            current244pin = 1;
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
                        cableMarkPlus.Position = cableLinePlus.StartPoint.Add(new Vector3d(-3, 0, 0));
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
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "3A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "+CABLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3M6/3." + currentPinAIAO[2];
                                                    currentPinAIAO[2]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3M6/4." + currentPinAIAO[3];
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.{1})", moduleNumber % 2 == 0 ? "1" : "2", sheetNum246);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPlus.TextString = string.Format("(-3A{0}:1/3.{1})", (moduleNumber + 4), currentSheet);
                                            }
                                            break;
                                        }
                                    case "LINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.{1})", moduleNumber % 2 == 0 ? "3" : "4", sheetNum246);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinus.TextString = string.Format("(-3A{0}:20/3.{1})", (moduleNumber + 4), currentSheet);
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "WA3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "AD":
                                        {
                                            if (firstAO)
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "AO";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    firstAO = false;
                                                }
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
                        cableMarkPlus.Position = cableLinePlus.StartPoint.Add(new Vector3d(-3, 0, 0));
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
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "3A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "+CABLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                if (moduleNumber % 2 == 0)
                                                {
                                                    attRef.TextString = "3M6/3." + currentPinAIAO[2];
                                                    currentPinAIAO[2]++;
                                                }
                                                else
                                                {
                                                    attRef.TextString = "3M6/4." + currentPinAIAO[3];
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.{1})", moduleNumber % 2 == 0 ? "1" : "2", sheetNum246);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPlus.TextString = string.Format("(-3A{0}:1/3.{1})", (moduleNumber+4), currentSheet); 
                                            }
                                            break;
                                        }
                                    case "LINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.6:{0}/3.{1})", moduleNumber % 2 == 0 ? "3" : "4", sheetNum246);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinus.TextString = string.Format("(-3A{0}:20/3.{1})", (moduleNumber + 4), currentSheet);
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "WA3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XT3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "SH":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "AD":
                                        {
                                            if (firstAI)
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "AI";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    firstAI = false;
                                                }
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3A" + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XP":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP3." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "8X":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDI1(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            #region entitiys
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block241.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            MText textUpPlus = new MText();
            textUpPlus.SetDatabaseDefaults();
            textUpPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textUpPlus.Location = block242.Add(new Vector3d(-3, -70, 0));
            textUpPlus.Rotation = 1.57;
            textUpPlus.Height = 2.5;
            textUpPlus.Attachment = AttachmentPoint.TopLeft;
            textUpPlus.Layer = "0";
            textUpPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textUpPlus);
            acTrans.AddNewlyCreatedDBObject(textUpPlus, true);

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
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkPlusUp.Position = block242.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

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
                        if (br.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props = br.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Видимость1" && !prop.ReadOnly)
                                {
                                    if (mainControl)
                                        prop.Value = values[0];
                                    else
                                        prop.Value = values[1];
                                    break;
                                }
                            }
                        }
                        BlockTableRecord btr = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                               
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G"+ currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "AD":
                                        {
                                            if (firstDI)
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "DI";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    firstDI = false;
                                                }
                                            }
                                            break;
                                        }
                                }
                                #endregion
                            }
                        }
                        changeLinks(acTrans);
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-1XR" + currentD.ToString() + "." + (moduleNumber + 4) + ":9:A2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:26/3." + currentSheet + ")";
                                                
                                                textUpPlus.Contents = "Клеммы питания\nреле DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textUpPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                                #endregion
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
            #region entitys
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block241.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            MText textUpPlus = new MText();
            textUpPlus.SetDatabaseDefaults();
            textUpPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textUpPlus.Location = block242.Add(new Vector3d(-3, -70, 0));
            textUpPlus.Rotation = 1.57;
            textUpPlus.Height = 2.5;
            textUpPlus.Attachment = AttachmentPoint.TopLeft;
            textUpPlus.Layer = "0";
            textUpPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textUpPlus);
            acTrans.AddNewlyCreatedDBObject(textUpPlus, true);

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
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkPlusUp.Position = cablePlusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp1.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1:2/3." + currentSheet + ")";
                                                textUpPlus.Contents = "Клеммы питания\nреле DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textUpPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) +".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                                textUpPlus.Contents = "Клеммы питания\nреле DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textUpPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
        }
        private static void insertDIGas(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            #region entitys
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block241.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            MText textUpPlus = new MText();
            textUpPlus.SetDatabaseDefaults();
            textUpPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textUpPlus.Location = block242.Add(new Vector3d(-3, -70, 0));
            textUpPlus.Rotation = 1.57;
            textUpPlus.Height = 2.5;
            textUpPlus.Attachment = AttachmentPoint.TopLeft;
            textUpPlus.Layer = "0";
            textUpPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textUpPlus);
            acTrans.AddNewlyCreatedDBObject(textUpPlus, true);

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
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkPlusUp.Position = cablePlusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp1.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1:2/3." + currentSheet + ")";
                                                textUpPlus.Contents = "Клеммы питания\nреле DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textUpPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
            }
            try
            {
                currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));
                ids = new ObjectIdCollection();
                filename = @"Data\GasDI.dwg";
                blockName = "GasDI";
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
                                    #region attributes
                                    switch (attDef.Tag)
                                    {
                                        case "TITLE":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = title;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "4A":
                                            {
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = string.Format("(-XT24.1:FU{0}/3.3)", moduleNumber + 1);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    linkPLusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                                    textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                    Extents3d ext = textDownPlus.GeometricExtents;
                                                    TypedValue[] filterlist = new TypedValue[2];
                                                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                    filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                    SelectionFilter filter = new SelectionFilter(filterlist);
                                                    Point3d point1 = ext.MinPoint;
                                                    Point3d point2 = ext.MaxPoint;
                                                    PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                    if (selRes.Status == PromptStatus.OK)
                                                    {
                                                        Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                        ent.Erase();
                                                    }
                                                }
                                                break;
                                            }
                                        case "UPCABLE-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = string.Format("(-XT24.2:{0}/3.3)", moduleNumber + 11);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    linkMinusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                                }
                                                break;
                                            }
                                        case "UPCABLE+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "1L" + (moduleNumber + 11) + "+";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    textDownPlus.Contents = "Сигнальные\nцепи DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                    Extents3d ext = textDownPlus.GeometricExtents;
                                                    TypedValue[] filterlist = new TypedValue[2];
                                                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                    filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                    SelectionFilter filter = new SelectionFilter(filterlist);
                                                    Point3d point1 = ext.MinPoint;
                                                    Point3d point2 = ext.MaxPoint;
                                                    PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                    if (selRes.Status == PromptStatus.OK)
                                                    {
                                                        Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                        ent.Erase();
                                                    }
                                                }
                                                break;
                                            }
                                        case "UPLINK+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = string.Format("(-XT24.2:FU{0}/3.3)", moduleNumber + 11);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    linkPLusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:2/3." + currentSheet + ")";
                                                    textUpPlus.Contents = "Клеммы питания\nреле DI\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                    Extents3d ext = textUpPlus.GeometricExtents;
                                                    TypedValue[] filterlist = new TypedValue[2];
                                                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                    filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                    SelectionFilter filter = new SelectionFilter(filterlist);
                                                    Point3d point1 = ext.MinPoint;
                                                    Point3d point2 = ext.MaxPoint;
                                                    PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                    if (selRes.Status == PromptStatus.OK)
                                                    {
                                                        Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                        ent.Erase();
                                                    }
                                                }
                                                break;
                                            }
                                        case "WA":
                                            {
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "3." + currentSheet;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "NUM":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentSheet.ToString();
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE3":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-3";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE4":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE5":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE6":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-6";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE7":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE8":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE9.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-9.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE9.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-9.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE10.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE10.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE11.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE11.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE12.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-12.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE12.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-12.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE13.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE13.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE14.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE14.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE15.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-15.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE15.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-15.2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE16.1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE16.2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16.1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R3":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".3";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD3":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".3";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD4":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".4";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD5":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".5";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2VD6":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2VD" + currentD.ToString() + "." + (moduleNumber + 4) + ".6";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE1+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI1P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI1M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE2+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI2P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI2M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE3+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI3P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE3-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI3M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE4+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI4P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE4-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI4M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE5+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI5P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE5-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI5M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE6+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI6P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE6-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDI6M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK1+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI1P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI1M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK2+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI2P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI2M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK3+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI3P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK3-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI3M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK4+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI4P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK4-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI4M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK5+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI5P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK5-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI5M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK6+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI6P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK6-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDI6M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                    #endregion
                                    }
                                }
                            }
                            lastSheetDI = currentSheet;
                            lastModNumDI = moduleNumber;
                            lastDNumDI = currentD;
                            currentSheet++;
                        }
                    }
                    else editor.WriteMessage("В файле не найден блок с именем \"{0}\"", blockName);
                }
            }
            catch
            {
                editor.WriteMessage("Ошибка отризовки DI с газом");
            }
        }

        private static void insertDO24V(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block243.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            #region down+
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block243;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block243 = block243.Add(new Vector3d(8, 0, 0));

            #region down-
            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block243;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            acCirc.Center = cableLineMinusDown.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCirc.Radius = 0.75;
            modSpace.AppendEntity(acCirc);
            acTrans.AddNewlyCreatedDBObject(acCirc, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCirc.ObjectId);

            Hatch acHatch = new Hatch();
            modSpace.AppendEntity(acHatch);
            acTrans.AddNewlyCreatedDBObject(acHatch, true);
            acHatch.SetDatabaseDefaults();
            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatch.Associative = true;
            acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatch.EvaluateHatch(true);
            acHatch.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusDown1 = new Line();
            cableBranchHMinusDown1.SetDatabaseDefaults();
            cableBranchHMinusDown1.Layer = "0";
            cableBranchHMinusDown1.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown1.StartPoint = acCirc.Center;
            cableBranchHMinusDown1.EndPoint = cableBranchHMinusDown1.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown1);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown1, true);

            Line cableBranchMinusDown1 = new Line();
            cableBranchMinusDown1.SetDatabaseDefaults();
            cableBranchMinusDown1.Layer = "0";
            cableBranchMinusDown1.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown1.StartPoint = cableBranchHMinusDown1.EndPoint;
            cableBranchMinusDown1.EndPoint = cableBranchMinusDown1.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown1);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown1, true);

            Line cableBranchHMinusDown2 = new Line();
            cableBranchHMinusDown2.SetDatabaseDefaults();
            cableBranchHMinusDown2.Layer = "0";
            cableBranchHMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown2.StartPoint = acCirc.Center;
            cableBranchHMinusDown2.EndPoint = cableBranchHMinusDown2.StartPoint.Add(new Vector3d(-4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown2, true);

            Line cableBranchMinusDown2 = new Line();
            cableBranchMinusDown2.SetDatabaseDefaults();
            cableBranchMinusDown2.Layer = "0";
            cableBranchMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown2.StartPoint = cableBranchHMinusDown2.EndPoint;
            cableBranchMinusDown2.EndPoint = cableBranchMinusDown2.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown2, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown1 = new DBText();
            linkMinusDown1.SetDatabaseDefaults();
            linkMinusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown1.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown1.Rotation = 1.57;
            linkMinusDown1.Height = 2.5;
            linkMinusDown1.WidthFactor = 0.7;
            linkMinusDown1.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown1.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown1.AlignmentPoint = linkMinusDown1.Position;
            linkMinusDown1.Layer = "0";
            modSpace.AppendEntity(linkMinusDown1);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown1, true);

            DBText linkMinusDown2 = new DBText();
            linkMinusDown2.SetDatabaseDefaults();
            linkMinusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown2.Position = cableBranchMinusDown1.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown2.Rotation = 1.57;
            linkMinusDown2.Height = 2.5;
            linkMinusDown2.WidthFactor = 0.7;
            linkMinusDown2.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown2.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown2.AlignmentPoint = linkMinusDown2.Position;
            linkMinusDown2.Layer = "0";
            modSpace.AppendEntity(linkMinusDown2);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown2, true);

            DBText linkMinusDown3 = new DBText();
            linkMinusDown3.SetDatabaseDefaults();
            linkMinusDown3.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown3.Position = cableBranchMinusDown2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown3.Rotation = 1.57;
            linkMinusDown3.Height = 2.5;
            linkMinusDown3.WidthFactor = 0.7;
            linkMinusDown3.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown3.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown3.AlignmentPoint = linkMinusDown3.Position;
            linkMinusDown3.Layer = "0";
            modSpace.AppendEntity(linkMinusDown3);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown3, true);
            #endregion

            block243 = block243.Add(new Vector3d(9.1624, 0, 0));

            #region up+
            Line cablePlusUp1 = new Line();
            cablePlusUp1.SetDatabaseDefaults();
            cablePlusUp1.Layer = "0";
            cablePlusUp1.Color = Color.FromRgb(255, 0, 0);
            cablePlusUp1.StartPoint = block244;
            cablePlusUp1.EndPoint = cablePlusUp1.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cablePlusUp1);
            acTrans.AddNewlyCreatedDBObject(cablePlusUp1, true);

            DBText cableMarkPlusUp1 = new DBText();
            cableMarkPlusUp1.SetDatabaseDefaults();
            cableMarkPlusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp1.Position = cablePlusUp1.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkPlusUp1.Rotation = 1.57;
            cableMarkPlusUp1.Height = 2.5;
            cableMarkPlusUp1.WidthFactor = 0.7;
            cableMarkPlusUp1.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp1.AlignmentPoint = cableMarkPlusUp1.Position;
            cableMarkPlusUp1.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp1);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp1, true);

            DBText linkPLusUp1 = new DBText();
            linkPLusUp1.SetDatabaseDefaults();
            linkPLusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp1.Position = cablePlusUp1.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp1.Rotation = 1.57;
            linkPLusUp1.Height = 2.5;
            linkPLusUp1.WidthFactor = 0.7;
            linkPLusUp1.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp1.AlignmentPoint = linkPLusUp1.Position;
            linkPLusUp1.Layer = "0";
            linkPLusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp1);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp1, true);

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            Line cableMinusUp1 = new Line();
            cableMinusUp1.SetDatabaseDefaults();
            cableMinusUp1.Layer = "0";
            cableMinusUp1.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp1.StartPoint = block244;
            cableMinusUp1.EndPoint = cableMinusUp1.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp1);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp1, true);

            DBText cableMarkMinusUp1 = new DBText();
            cableMarkMinusUp1.SetDatabaseDefaults();
            cableMarkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp1.Position = cableMinusUp1.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusUp1.Rotation = 1.57;
            cableMarkMinusUp1.Height = 2.5;
            cableMarkMinusUp1.WidthFactor = 0.7;
            cableMarkMinusUp1.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp1.AlignmentPoint = cableMarkMinusUp1.Position;
            cableMarkMinusUp1.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp1, true);

            DBText linkMinusUp1 = new DBText();
            linkMinusUp1.SetDatabaseDefaults();
            linkMinusUp1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp1.Position = cableMinusUp1.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp1.Rotation = 1.57;
            linkMinusUp1.Height = 2.5;
            linkMinusUp1.WidthFactor = 0.7;
            linkMinusUp1.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp1.AlignmentPoint = linkMinusUp1.Position;
            linkMinusUp1.Layer = "0";
            linkMinusUp1.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp1);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp1, true);
            #endregion

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            #region up-
            Line cablePlusUp2 = new Line();
            cablePlusUp2.SetDatabaseDefaults();
            cablePlusUp2.Layer = "0";
            cablePlusUp2.Color = Color.FromRgb(255, 0, 0);
            cablePlusUp2.StartPoint = block244;
            cablePlusUp2.EndPoint = cablePlusUp2.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cablePlusUp2);
            acTrans.AddNewlyCreatedDBObject(cablePlusUp2, true);

            DBText cableMarkPlusUp2 = new DBText();
            cableMarkPlusUp2.SetDatabaseDefaults();
            cableMarkPlusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp2.Position = cablePlusUp2.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkPlusUp2.Rotation = 1.57;
            cableMarkPlusUp2.Height = 2.5;
            cableMarkPlusUp2.WidthFactor = 0.7;
            cableMarkPlusUp2.Justify = AttachmentPoint.BottomRight;
            cableMarkPlusUp2.AlignmentPoint = cableMarkPlusUp2.Position;
            cableMarkPlusUp2.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkPlusUp2);
            acTrans.AddNewlyCreatedDBObject(cableMarkPlusUp2, true);

            DBText linkPLusUp2 = new DBText();
            linkPLusUp2.SetDatabaseDefaults();
            linkPLusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkPLusUp2.Position = cablePlusUp2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkPLusUp2.Rotation = 1.57;
            linkPLusUp2.Height = 2.5;
            linkPLusUp2.WidthFactor = 0.7;
            linkPLusUp2.Justify = AttachmentPoint.MiddleRight;
            linkPLusUp2.AlignmentPoint = linkPLusUp2.Position;
            linkPLusUp2.Layer = "0";
            linkPLusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkPLusUp2);
            acTrans.AddNewlyCreatedDBObject(linkPLusUp2, true);

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            Line cableMinusUp2 = new Line();
            cableMinusUp2.SetDatabaseDefaults();
            cableMinusUp2.Layer = "0";
            cableMinusUp2.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp2.StartPoint = block244;
            cableMinusUp2.EndPoint = cableMinusUp2.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp2);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp2, true);

            DBText cableMarkMinusUp2 = new DBText();
            cableMarkMinusUp2.SetDatabaseDefaults();
            cableMarkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp2.Position = cableMinusUp2.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusUp2.Rotation = 1.57;
            cableMarkMinusUp2.Height = 2.5;
            cableMarkMinusUp2.WidthFactor = 0.7;
            cableMarkMinusUp2.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusUp2.AlignmentPoint = cableMarkMinusUp2.Position;
            cableMarkMinusUp2.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusUp2, true);

            DBText linkMinusUp2 = new DBText();
            linkMinusUp2.SetDatabaseDefaults();
            linkMinusUp2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusUp2.Position = cableMinusUp2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusUp2.Rotation = 1.57;
            linkMinusUp2.Height = 2.5;
            linkMinusUp2.WidthFactor = 0.7;
            linkMinusUp2.Justify = AttachmentPoint.MiddleRight;
            linkMinusUp2.AlignmentPoint = linkMinusUp2.Position;
            linkMinusUp2.Layer = "0";
            linkMinusUp2.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(linkMinusUp2);
            acTrans.AddNewlyCreatedDBObject(linkMinusUp2, true);
            #endregion

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DO24V1.dwg";
            string blockName = "DO24V1";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2L" + (moduleNumber + 1 - diCount) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:FU{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":1/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Клеммы питания\nреле DO\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M-3/2." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp1.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:2.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1:1/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3-2L1+/1." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp1.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:1.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "AD":
                                        {
                                            if (firstDO)
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "DO";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    firstDO = false;
                                                }
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                        current244pin++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DO24V2.dwg";
            blockName = "DO24V2";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown2.TextString = "(-2G" + currentD.ToString() + "A" + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown3.TextString = "(-" + currentD.ToString() + "." + (moduleNumber + 4) + ":10/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M-3/2." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp2.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:2.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp2.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".2:1/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3-2L1+/1." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp2.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:1.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-2G" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "XT":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                        current244pin++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDO24Vhalf(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block243.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            #region down+
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block243;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block243 = block243.Add(new Vector3d(8, 0, 0));

            #region down-
            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block243;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            acCirc.Center = cableLineMinusDown.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCirc.Radius = 0.75;
            modSpace.AppendEntity(acCirc);
            acTrans.AddNewlyCreatedDBObject(acCirc, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCirc.ObjectId);

            Hatch acHatch = new Hatch();
            modSpace.AppendEntity(acHatch);
            acTrans.AddNewlyCreatedDBObject(acHatch, true);
            acHatch.SetDatabaseDefaults();
            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatch.Associative = true;
            acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatch.EvaluateHatch(true);
            acHatch.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusDown1 = new Line();
            cableBranchHMinusDown1.SetDatabaseDefaults();
            cableBranchHMinusDown1.Layer = "0";
            cableBranchHMinusDown1.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown1.StartPoint = acCirc.Center;
            cableBranchHMinusDown1.EndPoint = cableBranchHMinusDown1.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown1);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown1, true);

            Line cableBranchMinusDown1 = new Line();
            cableBranchMinusDown1.SetDatabaseDefaults();
            cableBranchMinusDown1.Layer = "0";
            cableBranchMinusDown1.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown1.StartPoint = cableBranchHMinusDown1.EndPoint;
            cableBranchMinusDown1.EndPoint = cableBranchMinusDown1.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown1);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown1, true);

            Line cableBranchHMinusDown2 = new Line();
            cableBranchHMinusDown2.SetDatabaseDefaults();
            cableBranchHMinusDown2.Layer = "0";
            cableBranchHMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown2.StartPoint = acCirc.Center;
            cableBranchHMinusDown2.EndPoint = cableBranchHMinusDown2.StartPoint.Add(new Vector3d(-4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown2, true);

            Line cableBranchMinusDown2 = new Line();
            cableBranchMinusDown2.SetDatabaseDefaults();
            cableBranchMinusDown2.Layer = "0";
            cableBranchMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown2.StartPoint = cableBranchHMinusDown2.EndPoint;
            cableBranchMinusDown2.EndPoint = cableBranchMinusDown2.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown2, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown1 = new DBText();
            linkMinusDown1.SetDatabaseDefaults();
            linkMinusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown1.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown1.Rotation = 1.57;
            linkMinusDown1.Height = 2.5;
            linkMinusDown1.WidthFactor = 0.7;
            linkMinusDown1.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown1.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown1.AlignmentPoint = linkMinusDown1.Position;
            linkMinusDown1.Layer = "0";
            modSpace.AppendEntity(linkMinusDown1);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown1, true);

            DBText linkMinusDown2 = new DBText();
            linkMinusDown2.SetDatabaseDefaults();
            linkMinusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown2.Position = cableBranchMinusDown1.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown2.Rotation = 1.57;
            linkMinusDown2.Height = 2.5;
            linkMinusDown2.WidthFactor = 0.7;
            linkMinusDown2.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown2.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown2.AlignmentPoint = linkMinusDown2.Position;
            linkMinusDown2.Layer = "0";
            modSpace.AppendEntity(linkMinusDown2);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown2, true);
            
            DBText linkMinusDown3 = new DBText();
            linkMinusDown3.SetDatabaseDefaults();
            linkMinusDown3.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown3.Position = cableBranchMinusDown2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown3.Rotation = 1.57;
            linkMinusDown3.Height = 2.5;
            linkMinusDown3.WidthFactor = 0.7;
            linkMinusDown3.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown3.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown3.AlignmentPoint = linkMinusDown3.Position;
            linkMinusDown3.Layer = "0";
            modSpace.AppendEntity(linkMinusDown3);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown3, true);
            #endregion

            block243 = block243.Add(new Vector3d(9.1624, 0, 0));

            #region up+
            Line cablePlusUp = new Line();
            cablePlusUp.SetDatabaseDefaults();
            cablePlusUp.Layer = "0";
            cablePlusUp.Color = Color.FromRgb(255, 0, 0);
            cablePlusUp.StartPoint = block244;
            cablePlusUp.EndPoint = cablePlusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cablePlusUp);
            acTrans.AddNewlyCreatedDBObject(cablePlusUp, true);

            DBText cableMarkPlusUp = new DBText();
            cableMarkPlusUp.SetDatabaseDefaults();
            cableMarkPlusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusUp.Position = cablePlusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            #region up-
            Line cableMinusUp = new Line();
            cableMinusUp.SetDatabaseDefaults();
            cableMinusUp.Layer = "0";
            cableMinusUp.Color = Color.FromRgb(0, 255, 255);
            cableMinusUp.StartPoint = block244;
            cableMinusUp.EndPoint = cableMinusUp.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableMinusUp);
            acTrans.AddNewlyCreatedDBObject(cableMinusUp, true);

            DBText cableMarkMinusUp = new DBText();
            cableMarkMinusUp.SetDatabaseDefaults();
            cableMarkMinusUp.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusUp.Position = cableMinusUp.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block244 = block244.Add(new Vector3d(3.4668, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DO24V1.dwg";
            string blockName = "DO24V1";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2L" + (moduleNumber + 1-diCount) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:FU{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":1/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Клеммы питания\nреле DO\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M-3/2." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:2.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusUp1.TextString = "(-XT" + currentD.ToString() + "." + (moduleNumber + 4) + ".1:1/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "UPCABLE+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3-2L1+/1." + current244pin;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusUp.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "UPLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.4:1.{0}/3.3)", current244pin);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusUp1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE3":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-3";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE6":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-6";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE9":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-9";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE12":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-12";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE15":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-15";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                        current244pin++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }

            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            filename = @"Data\DOM.dwg";
            blockName = "DOM";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown3.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":10/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDO1(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block243.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            #region down+
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block243;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block243 = block243.Add(new Vector3d(8, 0, 0));

            #region down-
            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block243;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            acCirc.Center = cableLineMinusDown.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCirc.Radius = 0.75;
            modSpace.AppendEntity(acCirc);
            acTrans.AddNewlyCreatedDBObject(acCirc, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCirc.ObjectId);

            Hatch acHatch = new Hatch();
            modSpace.AppendEntity(acHatch);
            acTrans.AddNewlyCreatedDBObject(acHatch, true);
            acHatch.SetDatabaseDefaults();
            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatch.Associative = true;
            acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatch.EvaluateHatch(true);
            acHatch.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusDown = new Line();
            cableBranchHMinusDown.SetDatabaseDefaults();
            cableBranchHMinusDown.Layer = "0";
            cableBranchHMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown.StartPoint = acCirc.Center;
            cableBranchHMinusDown.EndPoint = cableBranchHMinusDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown, true);

            Line cableBranchMinusDown = new Line();
            cableBranchMinusDown.SetDatabaseDefaults();
            cableBranchMinusDown.Layer = "0";
            cableBranchMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown.StartPoint = cableBranchHMinusDown.EndPoint;
            cableBranchMinusDown.EndPoint = cableBranchMinusDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown, true);

            Line cableBranchHMinusDown2 = new Line();
            cableBranchHMinusDown2.SetDatabaseDefaults();
            cableBranchHMinusDown2.Layer = "0";
            cableBranchHMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown2.StartPoint = acCirc.Center;
            cableBranchHMinusDown2.EndPoint = cableBranchHMinusDown2.StartPoint.Add(new Vector3d(-3, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown2, true);

            Line cableBranchMinusDown2 = new Line();
            cableBranchMinusDown2.SetDatabaseDefaults();
            cableBranchMinusDown2.Layer = "0";
            cableBranchMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown2.StartPoint = cableBranchHMinusDown2.EndPoint;
            cableBranchMinusDown2.EndPoint = cableBranchMinusDown2.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown2, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown1 = new DBText();
            linkMinusDown1.SetDatabaseDefaults();
            linkMinusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown1.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown1.Rotation = 1.57;
            linkMinusDown1.Height = 2.5;
            linkMinusDown1.WidthFactor = 0.7;
            linkMinusDown1.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown1.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown1.AlignmentPoint = linkMinusDown1.Position;
            linkMinusDown1.Layer = "0";
            modSpace.AppendEntity(linkMinusDown1);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown1, true);

            DBText linkMinusDown2 = new DBText();
            linkMinusDown2.SetDatabaseDefaults();
            linkMinusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown2.Position = cableBranchMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown2.Rotation = 1.57;
            linkMinusDown2.Height = 2.5;
            linkMinusDown2.WidthFactor = 0.7;
            linkMinusDown2.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown2.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown2.AlignmentPoint = linkMinusDown2.Position;
            linkMinusDown2.Layer = "0";
            modSpace.AppendEntity(linkMinusDown2);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown2, true);

            DBText linkMinusDown3 = new DBText();
            linkMinusDown3.SetDatabaseDefaults();
            linkMinusDown3.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown3.Position = cableBranchMinusDown2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown3.Rotation = 1.57;
            linkMinusDown3.Height = 2.5;
            linkMinusDown3.WidthFactor = 0.7;
            linkMinusDown3.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown3.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown3.AlignmentPoint = linkMinusDown3.Position;
            linkMinusDown3.Layer = "0";
            modSpace.AppendEntity(linkMinusDown3);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown3, true);
            #endregion

            block243 = block243.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DOP.dwg";
            string blockName = "DOP";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "R2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2L" + (moduleNumber + 1-diCount) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:FU{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":1/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Клеммы питания\nреле DO\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
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
            filename = @"Data\DO2.dwg";
            blockName = "DO2";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown2.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown3.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":10/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDO(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber)
        {
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block243.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            #region down+
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block243;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block243 = block243.Add(new Vector3d(8, 0, 0));

            #region down-
            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block243;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            acCirc.Center = cableLineMinusDown.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCirc.Radius = 0.75;
            modSpace.AppendEntity(acCirc);
            acTrans.AddNewlyCreatedDBObject(acCirc, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCirc.ObjectId);

            Hatch acHatch = new Hatch();
            modSpace.AppendEntity(acHatch);
            acTrans.AddNewlyCreatedDBObject(acHatch, true);
            acHatch.SetDatabaseDefaults();
            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatch.Associative = true;
            acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatch.EvaluateHatch(true);
            acHatch.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusDown = new Line();
            cableBranchHMinusDown.SetDatabaseDefaults();
            cableBranchHMinusDown.Layer = "0";
            cableBranchHMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown.StartPoint = acCirc.Center;
            cableBranchHMinusDown.EndPoint = cableBranchHMinusDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown, true);

            Line cableBranchMinusDown = new Line();
            cableBranchMinusDown.SetDatabaseDefaults();
            cableBranchMinusDown.Layer = "0";
            cableBranchMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown.StartPoint = cableBranchHMinusDown.EndPoint;
            cableBranchMinusDown.EndPoint = cableBranchMinusDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown, true);

            Line cableBranchHMinusDown2 = new Line();
            cableBranchHMinusDown2.SetDatabaseDefaults();
            cableBranchHMinusDown2.Layer = "0";
            cableBranchHMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown2.StartPoint = acCirc.Center;
            cableBranchHMinusDown2.EndPoint = cableBranchHMinusDown2.StartPoint.Add(new Vector3d(-3, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown2, true);

            Line cableBranchMinusDown2 = new Line();
            cableBranchMinusDown2.SetDatabaseDefaults();
            cableBranchMinusDown2.Layer = "0";
            cableBranchMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown2.StartPoint = cableBranchHMinusDown2.EndPoint;
            cableBranchMinusDown2.EndPoint = cableBranchMinusDown2.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown2, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown1 = new DBText();
            linkMinusDown1.SetDatabaseDefaults();
            linkMinusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown1.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown1.Rotation = 1.57;
            linkMinusDown1.Height = 2.5;
            linkMinusDown1.WidthFactor = 0.7;
            linkMinusDown1.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown1.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown1.AlignmentPoint = linkMinusDown1.Position;
            linkMinusDown1.Layer = "0";
            modSpace.AppendEntity(linkMinusDown1);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown1, true);

            DBText linkMinusDown2 = new DBText();
            linkMinusDown2.SetDatabaseDefaults();
            linkMinusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown2.Position = cableBranchMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown2.Rotation = 1.57;
            linkMinusDown2.Height = 2.5;
            linkMinusDown2.WidthFactor = 0.7;
            linkMinusDown2.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown2.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown2.AlignmentPoint = linkMinusDown2.Position;
            linkMinusDown2.Layer = "0";
            modSpace.AppendEntity(linkMinusDown2);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown2, true);

            DBText linkMinusDown3 = new DBText();
            linkMinusDown3.SetDatabaseDefaults();
            linkMinusDown3.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown3.Position = cableBranchMinusDown2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown3.Rotation = 1.57;
            linkMinusDown3.Height = 2.5;
            linkMinusDown3.WidthFactor = 0.7;
            linkMinusDown3.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown3.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown3.AlignmentPoint = linkMinusDown3.Position;
            linkMinusDown3.Layer = "0";
            modSpace.AppendEntity(linkMinusDown3);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown3, true);
            #endregion

            block243 = block243.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DO1.dwg";
            string blockName = "DO1";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2L" + (moduleNumber + 1-diCount) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:FU{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":1/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Клеммы питания\nреле DO\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":3/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
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
            filename = @"Data\DO2.dwg";
            blockName = "DO2";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1-diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "DOWNLINK1-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown2.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":10/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "DOWNLINK2-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown3.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
                                }
                            }
                        }
                        currentSheet++;
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void insertDOGas(Transaction acTrans, BlockTableRecord modSpace, Database acdb, int moduleNumber, bool gas)
        {
            MText textDownPlus = new MText();
            textDownPlus.SetDatabaseDefaults();
            textDownPlus.TextStyleId = tst["GOSTA-2.5-1"];
            textDownPlus.Location = block243.Add(new Vector3d(-3, -70, 0));
            textDownPlus.Rotation = 1.57;
            textDownPlus.Height = 2.5;
            textDownPlus.Attachment = AttachmentPoint.TopLeft;
            textDownPlus.Layer = "0";
            textDownPlus.Color = Color.FromRgb(255, 255, 255);
            modSpace.AppendEntity(textDownPlus);
            acTrans.AddNewlyCreatedDBObject(textDownPlus, true);

            #region down+
            Line cableDownPlus = new Line();
            cableDownPlus.SetDatabaseDefaults();
            cableDownPlus.Layer = "0";
            cableDownPlus.Color = Color.FromRgb(255, 0, 0);
            cableDownPlus.StartPoint = block243;
            cableDownPlus.EndPoint = cableDownPlus.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableDownPlus);
            acTrans.AddNewlyCreatedDBObject(cableDownPlus, true);

            DBText cableMarkPlusDown = new DBText();
            cableMarkPlusDown.SetDatabaseDefaults();
            cableMarkPlusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkPlusDown.Position = cableDownPlus.StartPoint.Add(new Vector3d(0, -1, 0));
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
            #endregion

            block243 = block243.Add(new Vector3d(8, 0, 0));

            #region down-
            Line cableLineMinusDown = new Line();
            cableLineMinusDown.SetDatabaseDefaults();
            cableLineMinusDown.Layer = "0";
            cableLineMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableLineMinusDown.StartPoint = block243;
            cableLineMinusDown.EndPoint = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -16.6269, 0));
            modSpace.AppendEntity(cableLineMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableLineMinusDown, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            acCirc.Center = cableLineMinusDown.EndPoint.Add(new Vector3d(0, 3.9489, 0));
            acCirc.Radius = 0.75;
            modSpace.AppendEntity(acCirc);
            acTrans.AddNewlyCreatedDBObject(acCirc, true);

            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
            acObjIdColl.Add(acCirc.ObjectId);

            Hatch acHatch = new Hatch();
            modSpace.AppendEntity(acHatch);
            acTrans.AddNewlyCreatedDBObject(acHatch, true);
            acHatch.SetDatabaseDefaults();
            acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatch.Associative = true;
            acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
            acHatch.EvaluateHatch(true);
            acHatch.Color = Color.FromRgb(255, 255, 255);

            Line cableBranchHMinusDown = new Line();
            cableBranchHMinusDown.SetDatabaseDefaults();
            cableBranchHMinusDown.Layer = "0";
            cableBranchHMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown.StartPoint = acCirc.Center;
            cableBranchHMinusDown.EndPoint = cableBranchHMinusDown.StartPoint.Add(new Vector3d(4, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown, true);

            Line cableBranchMinusDown = new Line();
            cableBranchMinusDown.SetDatabaseDefaults();
            cableBranchMinusDown.Layer = "0";
            cableBranchMinusDown.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown.StartPoint = cableBranchHMinusDown.EndPoint;
            cableBranchMinusDown.EndPoint = cableBranchMinusDown.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown, true);

            Line cableBranchHMinusDown2 = new Line();
            cableBranchHMinusDown2.SetDatabaseDefaults();
            cableBranchHMinusDown2.Layer = "0";
            cableBranchHMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchHMinusDown2.StartPoint = acCirc.Center;
            cableBranchHMinusDown2.EndPoint = cableBranchHMinusDown2.StartPoint.Add(new Vector3d(-3, 0, 0));
            modSpace.AppendEntity(cableBranchHMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchHMinusDown2, true);

            Line cableBranchMinusDown2 = new Line();
            cableBranchMinusDown2.SetDatabaseDefaults();
            cableBranchMinusDown2.Layer = "0";
            cableBranchMinusDown2.Color = Color.FromRgb(0, 255, 255);
            cableBranchMinusDown2.StartPoint = cableBranchHMinusDown2.EndPoint;
            cableBranchMinusDown2.EndPoint = cableBranchMinusDown2.StartPoint.Add(new Vector3d(0, -3.9489, 0));
            modSpace.AppendEntity(cableBranchMinusDown2);
            acTrans.AddNewlyCreatedDBObject(cableBranchMinusDown2, true);

            DBText cableMarkMinusDown = new DBText();
            cableMarkMinusDown.SetDatabaseDefaults();
            cableMarkMinusDown.TextStyleId = tst["GOSTA-2.5-1"];
            cableMarkMinusDown.Position = cableLineMinusDown.StartPoint.Add(new Vector3d(0, -1, 0));
            cableMarkMinusDown.Rotation = 1.57;
            cableMarkMinusDown.Height = 2.5;
            cableMarkMinusDown.WidthFactor = 0.7;
            cableMarkMinusDown.Justify = AttachmentPoint.BottomRight;
            cableMarkMinusDown.AlignmentPoint = cableMarkMinusDown.Position;
            cableMarkMinusDown.Layer = "КИА_МАРКИРОВКА";
            modSpace.AppendEntity(cableMarkMinusDown);
            acTrans.AddNewlyCreatedDBObject(cableMarkMinusDown, true);

            DBText linkMinusDown1 = new DBText();
            linkMinusDown1.SetDatabaseDefaults();
            linkMinusDown1.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown1.Position = cableLineMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown1.Rotation = 1.57;
            linkMinusDown1.Height = 2.5;
            linkMinusDown1.WidthFactor = 0.7;
            linkMinusDown1.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown1.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown1.AlignmentPoint = linkMinusDown1.Position;
            linkMinusDown1.Layer = "0";
            modSpace.AppendEntity(linkMinusDown1);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown1, true);

            DBText linkMinusDown2 = new DBText();
            linkMinusDown2.SetDatabaseDefaults();
            linkMinusDown2.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown2.Position = cableBranchMinusDown.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown2.Rotation = 1.57;
            linkMinusDown2.Height = 2.5;
            linkMinusDown2.WidthFactor = 0.7;
            linkMinusDown2.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown2.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown2.AlignmentPoint = linkMinusDown2.Position;
            linkMinusDown2.Layer = "0";
            modSpace.AppendEntity(linkMinusDown2);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown2, true);

            DBText linkMinusDown3 = new DBText();
            linkMinusDown3.SetDatabaseDefaults();
            linkMinusDown3.TextStyleId = tst["GOSTA-2.5-1"];
            linkMinusDown3.Position = cableBranchMinusDown2.EndPoint.Add(new Vector3d(0, -1, 0));
            linkMinusDown3.Rotation = 1.57;
            linkMinusDown3.Height = 2.5;
            linkMinusDown3.WidthFactor = 0.7;
            linkMinusDown3.Color = Color.FromRgb(255, 255, 255);
            linkMinusDown3.Justify = AttachmentPoint.MiddleRight;
            linkMinusDown3.AlignmentPoint = linkMinusDown3.Position;
            linkMinusDown3.Layer = "0";
            modSpace.AppendEntity(linkMinusDown3);
            acTrans.AddNewlyCreatedDBObject(linkMinusDown3, true);
            #endregion

            block243 = block243.Add(new Vector3d(9.1624, 0, 0));

            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\DO1.dwg";
            string blockName = "DO1";
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
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "TITLE":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = title;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "4A":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2L" + (moduleNumber + 1 - diCount) + "+";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkPlusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNCABLE-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                cableMarkMinusDown.TextString = attRef.TextString;
                                            }
                                            break;
                                        }
                                    case "DOWNLINK+":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:FU{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkPLusDown1.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":1/3." + currentSheet + ")";
                                                textDownPlus.Contents = "Клеммы питания\nреле DO\nМодуль " + currentD + "A" + (moduleNumber + 4);
                                                Extents3d ext = textDownPlus.GeometricExtents;
                                                TypedValue[] filterlist = new TypedValue[2];
                                                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                                filterlist[1] = new TypedValue((int)DxfCode.Text, "Резерв");
                                                SelectionFilter filter = new SelectionFilter(filterlist);
                                                Point3d point1 = ext.MinPoint;
                                                Point3d point2 = ext.MaxPoint;
                                                PromptSelectionResult selRes = editor.SelectWindow(point1, point2);
                                                if (selRes.Status == PromptStatus.OK)
                                                {
                                                    Entity ent = (Entity)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                                                    ent.Erase();
                                                }
                                            }
                                            break;
                                        }
                                    case "DOWNLINK-":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                linkMinusDown1.TextString = "(-1G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                            }
                                            break;
                                        }
                                    case "WA":
                                        {
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = "3." + currentSheet;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "NUM":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheet.ToString();
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE1":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-1";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE2":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-2";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE4":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-4";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE5":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-5";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE7":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-7";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE8":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-8";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE10":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-10";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE11":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-11";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE13":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-13";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE14":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-14";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE16":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-16";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE17":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-17";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE19":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-19";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE20":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-20";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE22":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-22";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE23":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-23";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE25":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-25";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE26":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-26";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE28":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-28";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE29":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-29";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE31":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-31";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE32":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-32";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE34":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-34";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE35":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-35";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE37":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-37";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE38":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-38";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE40":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-40";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE41":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-41";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE43":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-43";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE44":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-44";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE46":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-46";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                    case "CABLE47":
                                        {
                                            AttributeReference attRef = new AttributeReference();
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".1-47";
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                #endregion
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
            if (gas)
            {
                filename = @"Data\GasDO.dwg";
                blockName = "GasDO";
            }
            else
            {
                filename = @"Data\NonGasDO.dwg";
                blockName = "NonGasDO";
            }
            try
            {
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
                                    #region attributes
                                    switch (attDef.Tag)
                                    {
                                        case "TITLE":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = title;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "4A":
                                            {
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "-XP" + currentD.ToString() + "." + (moduleNumber + 4);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "DOWNCABLE1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "DOWNCABLE2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2M" + (moduleNumber + 1 - diCount);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "DOWNLINK1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    linkMinusDown2.TextString = "(-" + currentD.ToString() + "A" + (moduleNumber + 4) + ":10/3." + currentSheet + ")";
                                                }
                                                break;
                                            }
                                        case "DOWNLINK2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = string.Format("(-XT24.3:{0}/3.4)", moduleNumber + 1 - diCount);
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    linkMinusDown3.TextString = "(-2G" + currentD.ToString() + "." + (moduleNumber + 4) + ":2/3." + currentSheet + ")";
                                                }
                                                break;
                                            }
                                        case "WA":
                                            {
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
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
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "3." + currentSheet;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "NUM":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentSheet.ToString();
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE4":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-4";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE5":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-5";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE7":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-7";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE8":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-8";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE10":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-10";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE11":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-11";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE13":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-13";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE14":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-14";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE16":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-16";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE17":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-17";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE19":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-19";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE20":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-20";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE22":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-22";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE23":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-23";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE25":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-25";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE27":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-27";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE28":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-28";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE30":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-30";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE31":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-31";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE33":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-33";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE34":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-34";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE36":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-36";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE37":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-37";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE39":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-39";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE40":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-40";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE42":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-42";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE43":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-43";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE45":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-45";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE46":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-46";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "CABLE48":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = currentD.ToString() + "." + (moduleNumber + 4) + ".2-48";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R1":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".1";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R2":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".2";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R3":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".3";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R4":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".4";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R5":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".5";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R6":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".6";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R7":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".7";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R8":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".8";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R9":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".9";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R10":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".10";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R11":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".11";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R12":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".12";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R13":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".13";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R14":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".14";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R15":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".15";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "2R16":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    attRef.TextString = "2R" + currentD.ToString() + "." + (moduleNumber + 4) + ".16";
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE1+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO1P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO1M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE2+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO2P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO2M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE3+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO3P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE3-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO3M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE4+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO4P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASCABLE4-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasCableDO4M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK1+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO1P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK1-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO1M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK2+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO2P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK2-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO2M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK3+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO3P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK3-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO3M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK4+":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO4P = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                        case "GASLINK4-":
                                            {
                                                AttributeReference attRef = new AttributeReference();
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    gasLinkDO4M = attRef;
                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                                break;
                                            }
                                    #endregion
                                    }
                                }
                            }
                            lastSheetDO = currentSheet;
                            lastModNumDO = moduleNumber;
                            lastDNumDO = currentD;
                            currentSheet++;
                        }
                    }
                    else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
                }
            }
            catch
            {
                editor.WriteMessage("Ошибка отризовки DO " + (gas?"с газом":"без газа"));
            }
            
            currentPoint = currentPoint.Add(new Vector3d(594, 0, 0));

            ids = new ObjectIdCollection();
            if (gas)
            {
                try
                {
                    filename = @"Data\Gas.dwg";
                    blockName = "Gas";
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
                                        #region attributes
                                        switch (attDef.Tag)
                                        {
                                            case "TITLE":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = title;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE1+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-1+";
                                                        gasCableDI1P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE1-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-1-";
                                                        gasCableDI1M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE2+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-2+";
                                                        gasCableDI2P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE2-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-2-";
                                                        gasCableDI2M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE3+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-1+";
                                                        gasCableDI3P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE3-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-1-";
                                                        gasCableDI3M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE4+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-4+";
                                                        gasCableDI4P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE4-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-4-";
                                                        gasCableDI4M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE5+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-04-2+";
                                                        gasCableDI5P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE5-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-04-2-";
                                                        gasCableDI5M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE6+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-04-5+";
                                                        gasCableDI6P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE6-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-04-5-";
                                                        gasCableDI6M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE7+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-3+";
                                                        gasCableDO1P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE7-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-3-";
                                                        gasCableDO1M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE8+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-6+";
                                                        gasCableDO2P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE8-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-6-";
                                                        gasCableDO2M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE9+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-7+";
                                                        gasCableDO3P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE9-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-7-";
                                                        gasCableDO3M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE10+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-8+";
                                                        gasCableDO4P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE10-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-8-";
                                                        gasCableDO4M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK1+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:25/3." + lastSheetDI + ")";
                                                        gasLinkDI1P.TextString = "(-XT:TB-01:1+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK1-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:26/3." + lastSheetDI + ")";
                                                        gasLinkDI1M.TextString = "(-XT:TB-01:1-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK2+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:28/3." + lastSheetDI + ")";
                                                        gasLinkDI2P.TextString = "(-XT:TB-01:2+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK2-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:29/3." + lastSheetDI + ")";
                                                        gasLinkDI2M.TextString = "(-XT:TB-01:2-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK3+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:34/3." + lastSheetDI + ")";
                                                        gasLinkDI3P.TextString = "(-XT:TB-02:1+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK3-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:35/3." + lastSheetDI + ")";
                                                        gasLinkDI3M.TextString = "(-XT:TB-02:1-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK4+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:37/3." + lastSheetDI + ")";
                                                        gasLinkDI4P.TextString = "(-XT:TB-02:4+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK4-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:38/3." + lastSheetDI + ")";
                                                        gasLinkDI4M.TextString = "(-XT:TB-02:4-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK5+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:40/3." + lastSheetDI + ")";
                                                        gasLinkDI5P.TextString = "(-XT:TB-04:2+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK5-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:41/3." + lastSheetDI + ")";
                                                        gasLinkDI5M.TextString = "(-XT:TB-04:2-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK6+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:43/3." + lastSheetDI + ")";
                                                        gasLinkDI6P.TextString = "(-XT:TB-04:5+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK6-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:44/3." + lastSheetDI + ")";
                                                        gasLinkDI6M.TextString = "(-XT:TB-04:5-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK7+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:25/3." + lastSheetDO + ")";
                                                        gasLinkDO1P.TextString = "(-XT:TB-03:3+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK7-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:26/3." + lastSheetDO + ")";
                                                        gasLinkDO1M.TextString = "(-XT:TB-03:3-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK8+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:28/3." + lastSheetDO + ")";
                                                        gasLinkDO2P.TextString = "(-XT:TB-03:6+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK8-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:29/3." + lastSheetDO + ")";
                                                        gasLinkDO2M.TextString = "(-XT:TB-03:6-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK9+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:31/3." + lastSheetDO + ")";
                                                        gasLinkDO3P.TextString = "(-XT:TB-03:7+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK9-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:32/3." + lastSheetDO + ")";
                                                        gasLinkDO3M.TextString = "(-XT:TB-03:7-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK10+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:34/3." + lastSheetDO + ")";
                                                        gasLinkDO4P.TextString = "(-XT:TB-03:8+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK10-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:35/3." + lastSheetDO + ")";
                                                        gasLinkDO4M.TextString = "(-XT:TB-03:8-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                        #endregion
                                        }
                                    }
                                }
                                currentSheet++;
                            }
                        }
                        else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
                    }
                }
                catch
                {
                    editor.WriteMessage("Ошибка отризовки листа с газом");
                }
            }
            else
            {
                try
                {
                    filename = @"Data\NonGas.dwg";
                    blockName = "NonGas";
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
                                        #region attributes
                                        switch (attDef.Tag)
                                        {
                                            case "TITLE":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = title;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE1+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-1+";
                                                        gasCableDI1P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE1-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-1-";
                                                        gasCableDI1M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE2+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-2+";
                                                        gasCableDI2P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE2-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-01-2-";
                                                        gasCableDI2M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE3+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-1+";
                                                        gasCableDI3P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE3-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-1-";
                                                        gasCableDI3M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE4+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-3+";
                                                        gasCableDI4P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE4-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-3-";
                                                        gasCableDI4M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE5+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-4+";
                                                        gasCableDI5P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE5-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-4-";
                                                        gasCableDI5M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE6+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-5+";
                                                        gasCableDI6P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE6-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-02-5-";
                                                        gasCableDI6M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE7+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-2+";
                                                        gasCableDO1P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE7-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-2-";
                                                        gasCableDO1M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE8+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-6+";
                                                        gasCableDO2P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE8-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-6-";
                                                        gasCableDO2M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE9+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-7+";
                                                        gasCableDO3P.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASCABLE9-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "TB-03-7-";
                                                        gasCableDO3M.TextString = attRef.TextString;
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK1+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:25/3." + lastSheetDI + ")";
                                                        gasLinkDI1P.TextString = "(-XT:TB-01:1+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK1-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:26/3." + lastSheetDI + ")";
                                                        gasLinkDI1M.TextString = "(-XT:TB-01:1-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK2+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:28/3." + lastSheetDI + ")";
                                                        gasLinkDI2P.TextString = "(-XT:TB-01:2+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK2-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:29/3." + lastSheetDI + ")";
                                                        gasLinkDI2M.TextString = "(-XT:TB-01:2-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK3+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:34/3." + lastSheetDI + ")";
                                                        gasLinkDI3P.TextString = "(-XT:TB-02:1+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK3-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:35/3." + lastSheetDI + ")";
                                                        gasLinkDI3M.TextString = "(-XT:TB-02:1-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK4+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:37/3." + lastSheetDI + ")";
                                                        gasLinkDI4P.TextString = "(-XT:TB-02:3+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK4-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:38/3." + lastSheetDI + ")";
                                                        gasLinkDI4M.TextString = "(-XT:TB-02:3-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK5+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:40/3." + lastSheetDI + ")";
                                                        gasLinkDI5P.TextString = "(-XT:TB-02:4+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK5-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:41/3." + lastSheetDI + ")";
                                                        gasLinkDI5M.TextString = "(-XT:TB-02:4-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK6+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:43/3." + lastSheetDI + ")";
                                                        gasLinkDI6P.TextString = "(-XT:TB-02:5+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK6-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDI + "." + lastModNumDI + ".2:44/3." + lastSheetDI + ")";
                                                        gasLinkDI6M.TextString = "(-XT:TB-02:5-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK7+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:25/3." + lastSheetDO + ")";
                                                        gasLinkDO1P.TextString = "(-XT:TB-03:2+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK7-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:26/3." + lastSheetDO + ")";
                                                        gasLinkDO1M.TextString = "(-XT:TB-03:2-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK8+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:28/3." + lastSheetDO + ")";
                                                        gasLinkDO2P.TextString = "(-XT:TB-03:6+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK8-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:29/3." + lastSheetDO + ")";
                                                        gasLinkDO2M.TextString = "(-XT:TB-03:6-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK9+":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:31/3." + lastSheetDO + ")";
                                                        gasLinkDO3P.TextString = "(-XT:TB-03:7+/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                            case "GASLINK9-":
                                                {
                                                    AttributeReference attRef = new AttributeReference();
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                        attRef.TextString = "(-XT" + lastDNumDO + "." + lastModNumDO + ".2:32/3." + lastSheetDO + ")";
                                                        gasLinkDO3M.TextString = "(-XT:TB-03:7-/3." + currentSheet + ")";
                                                        br.AttributeCollection.AppendAttribute(attRef);
                                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                    break;
                                                }
                                        #endregion
                               
                                        }
                                    }
                                }
                                currentSheet++;
                            }
                        }
                        else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
                    }
                }
                catch
                {
                    editor.WriteMessage("Ошибка отризовки листа без газа");
                }
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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
                                            AttributeReference attRef = new AttributeReference();
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

        private static void insertSpec(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename = @"Data\Specification.dwg";
            string blockName = "Specification";
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
                                 #region attributes
                                 switch (attDef.Tag)
                                 {
                                     case "TITLE1":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE2":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE3":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE4":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE5":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE6":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE7":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "TITLE8":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = title;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH1":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + currentSheet;
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH2":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet+1);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH3":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 2);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH4":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 3);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH5":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 4);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH6":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 5);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH7":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 6);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                     case "SH8":
                                         {
                                             AttributeReference attRef = new AttributeReference();
                                             {
                                                 attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                 attRef.TextString = "3." + (currentSheet + 7);
                                                 br.AttributeCollection.AppendAttribute(attRef);
                                                 acTrans.AddNewlyCreatedDBObject(attRef, true);
                                             }
                                             break;
                                         }
                                 #endregion
                                 }
                             }
                         }
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static Table insertTable(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint)
        {
            Table table = new Table();
            table.Position = insertPoint.Add(new Vector3d(132, -26, 0));
            table.SetSize(8, 1);
            table.TableStyle = acdb.Tablestyle;
            table.Rotation = 1.5708;

            Matrix3d curUCSMatrix = acDoc.Editor.CurrentUserCoordinateSystem;
            CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;
            table.TransformBy(Matrix3d.Rotation(1.5708,
                                                 curUCS.Zaxis,
                                                 table.Position));
            table.Position = table.Position.Add(new Vector3d(0, -137, 0));

            table.SetBreakHeight(0, 191);
            table.SetBreakSpacing(28);
            table.SetBreakOffset(0, new Vector3d(28, 0, 0));
            table.BreakFlowDirection = TableBreakFlowDirection.DownOrUp;
            table.BreakEnabled = true;

            table.SetTextHeight(0, 0, 2.5);
            table.Cells[0, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[0, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(0, 0, CellAlignment.MiddleCenter);
            table.Rows[0].IsMergeAllEnabled = false;

            table.SetTextHeight(1, 0, 2.5);
            table.Cells[1, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[1, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(1, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(2, 0, 2.5);
            table.Cells[2, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[2, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(2, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(3, 0, 2.5);
            table.Cells[3, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[3, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(3, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(4, 0, 2.5);
            table.Cells[4, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[4, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(4, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(5, 0, 2.5);
            table.Cells[5, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[5, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(5, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(6, 0, 2.5);
            table.Cells[6, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[6, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(6, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(7, 0, 2.5);
            table.Cells[7, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[7, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(7, 0, CellAlignment.MiddleCenter);

            table.Columns[0].Width = 137;
            table.SetRowHeight(0, 40);
            table.SetRowHeight(1, 40);
            table.SetRowHeight(2, 40);
            table.SetRowHeight(3, 40);
            table.SetRowHeight(4, 40);
            table.SetRowHeight(5, 40);
            table.SetRowHeight(6, 40);
            table.SetRowHeight(7, 40);

            table.GenerateLayout();
            modSpace.AppendEntity(table);
            acTrans.AddNewlyCreatedDBObject(table, true);
            return table;
        }

        private static void addRowsToTable(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d insertPoint, Table table)
        {
            table.SetSize(table.Rows.Count + 8, 1);
            table.BreakOptions = TableBreakOptions.AllowManualPositions;
            table.SetBreakHeight(2, 191);
            table.SetBreakSpacing(28);
            table.SetBreakOffset(2, new Vector3d(table.Height + 246, 0, 0));
            table.BreakFlowDirection = TableBreakFlowDirection.DownOrUp;
            table.BreakEnabled = true;

            table.SetTextHeight(table.Rows.Count - 8, 0, 2.5);
            table.Cells[table.Rows.Count - 8, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 8, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 8, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 7, 0, 2.5);
            table.Cells[table.Rows.Count - 7, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 7, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 7, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 6, 0, 2.5);
            table.Cells[table.Rows.Count - 6, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 6, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 6, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 5, 0, 2.5);
            table.Cells[table.Rows.Count - 5, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 5, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 5, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 4, 0, 2.5);
            table.Cells[table.Rows.Count - 4, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 4, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 4, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 3, 0, 2.5);
            table.Cells[table.Rows.Count - 3, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 3, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 3, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 2, 0, 2.5);
            table.Cells[table.Rows.Count - 2, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 2, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 2, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(table.Rows.Count - 1, 0, 2.5);
            table.Cells[table.Rows.Count - 1, 0].TextString = "";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[table.Rows.Count - 1, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(table.Rows.Count - 1, 0, CellAlignment.MiddleCenter);

            table.Columns[0].Width = 137;
            table.SetRowHeight(table.Rows.Count - 8, 40);
            table.SetRowHeight(table.Rows.Count - 7, 40);
            table.SetRowHeight(table.Rows.Count - 6, 40);
            table.SetRowHeight(table.Rows.Count - 5, 40);
            table.SetRowHeight(table.Rows.Count - 4, 40);
            table.SetRowHeight(table.Rows.Count - 3, 40);
            table.SetRowHeight(table.Rows.Count - 2, 40);
            table.SetRowHeight(table.Rows.Count - 1, 40);

            table.GenerateLayout();
        }

        private static void setPoints(Point3d startPoint)
        {
            link1 = startPoint.Add(new Vector3d(1035.5134, -200.7757, 0));
            link2 = link1.Add(new Vector3d(84.3448, 22.5574, 0));
            link3 = link2.Add(new Vector3d(18.9336, -22.5574, 0));
            block241 = link3.Add(new Vector3d(155.0528, -53.042, 0));
            block242 = block241.Add(new Vector3d(229.5485, 0, 0));
            link4 = block242.Add(new Vector3d(-92.2837, 52.978, 0));
            block243 = link4.Add(new Vector3d(429.3387, -52.978, 0));
            block244 = block243.Add(new Vector3d(204.8261, 0, 0));
            link5 = block244.Add(new Vector3d(-67.5613, 52.978, 0));
            link6 = link5.Add(new Vector3d(286.8018, 0, 0));
            if (mainControl)
            {
                link7 = link6.Add(new Vector3d(339.4296, -96.3883, 0));
                link8 = link7.Add(new Vector3d(149.9289, 35.7228, 0));
                ets.Add(link8.Add(new Vector3d(315.4449, 225.3946, 0)));
            }
            else ets.Add(link6.Add(new Vector3d(210.8035, 164.7291, 0)));
            ets.Add(ets[0].Add(new Vector3d(0, -145.3992, 0)));
            ets.Add(ets[1].Add(new Vector3d(0, -145.3992, 0)));
            block246 = ets[1].Add(new Vector3d(319.6824, 148.6242, 0));
            currentPoint = block246.Add(new Vector3d(1342.9993, 32.8856, 0));
        }

        private static void changeLinks(Transaction acTrans)
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point1 = link1;
            Point3d point2 = point1.Add(new Vector3d(0, 5, 0));
            Point3dCollection points = new Point3dCollection(new Point3d[] { point1, point2 });
            PromptSelectionResult selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №1\n");
            }
            else
            {
                DBText link1Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link1Text.TextString = string.Format("(-1XR4.4:11:A1/3.{0})", currentSheet);
            }

            filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            filter = new SelectionFilter(filterlist);
            point1 = link2;
            point2 = point1.Add(new Vector3d(0, -5, 0));
            points = new Point3dCollection(new Point3d[] { point1, point2 });
            selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №2\n");
            }
            else
            {
                DBText link2Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link2Text.TextString = string.Format("(-KV2:11/3.{0})", currentSheet);
            }

            filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            filter = new SelectionFilter(filterlist);
            point1 = link3;
            point2 = point1.Add(new Vector3d(0, 5, 0));
            points = new Point3dCollection(new Point3d[] { point1, point2 });
            selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №3\n");
            }
            else
            {
                DBText link3Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link3Text.TextString = string.Format("(-1XR4.4:10:A1/3.{0})", currentSheet);
            }

            filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            filter = new SelectionFilter(filterlist);
            point1 = link4;
            point2 = point1.Add(new Vector3d(5, 0, 0));
            points = new Point3dCollection(new Point3d[] { point1, point2 });
            selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №4\n");
            }
            else
            {
                DBText link4Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link4Text.TextString = string.Format("(-1XR4.4:12:A1/3.{0})", currentSheet);
            }

            filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            filter = new SelectionFilter(filterlist);
            point1 = link5;
            point2 = point1.Add(new Vector3d(5, 0, 0));
            points = new Point3dCollection(new Point3d[] { point1, point2 });
            selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №5\n");
            }
            else
            {
                DBText link5Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link5Text.TextString = string.Format("(-1XR4.4:13:A1/3.{0})", currentSheet);
            }

            filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
            filter = new SelectionFilter(filterlist);
            point1 = link6;
            point2 = point1.Add(new Vector3d(5, 0, 0));
            points = new Point3dCollection(new Point3d[] { point1, point2 });
            selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе найдена ссылка №6\n");
            }
            else
            {
                DBText link6Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                link6Text.TextString = string.Format("(-1XR4.4:14:A1/3.{0})", currentSheet);
            }

            if (mainControl)
            {
                filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
                filter = new SelectionFilter(filterlist);
                point1 = link7;
                point2 = point1.Add(new Vector3d(5, 0, 0));
                points = new Point3dCollection(new Point3d[] { point1, point2 });
                selRes = editor.SelectFence(points, filter);
                if (selRes.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("\nНе найдена ссылка №7\n");
                }
                else
                {
                    DBText link7Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                    link7Text.TextString = string.Format("(-1XR4.4:15:A1/3.{0})", currentSheet);
                }

                filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Временный");
                filter = new SelectionFilter(filterlist);
                point1 = link8;
                point2 = point1.Add(new Vector3d(5, 0, 0));
                points = new Point3dCollection(new Point3d[] { point1, point2 });
                selRes = editor.SelectFence(points, filter);
                if (selRes.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("\nНе найдена ссылка №8\n");
                }
                else
                {
                    DBText link8Text = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForWrite);
                    link8Text.TextString = string.Format("(-1XR4.4:16:A1/3.{0})", currentSheet);
                }
            }
        }
    }
}