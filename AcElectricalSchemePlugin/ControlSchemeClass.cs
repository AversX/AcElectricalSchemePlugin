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
        private static int aiCount = 0;
        private static int aoCount = 0;
        private static int diCount = 0;
        private static int doCount = 0;

        static public void fixLinks()
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            PromptIntegerResult AI = editor.GetInteger("Введите количество модулей AI: ");
            if (AI.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else aiCount = AI.Value;

            PromptIntegerResult AO = editor.GetInteger("Введите количество модулей AO: ");
            if (AO.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else aoCount = AO.Value;

            PromptIntegerResult DI = editor.GetInteger("Введите количество модулей DI: ");
            if (DI.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else diCount = DI.Value;

            PromptIntegerResult DO = editor.GetInteger("Введите количество модулей DO: ");
            if (DO.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else doCount = DO.Value;

            PromptEntityResult aiao = acDoc.Editor.GetEntity("Выберите крайнюю правую линию листа с клеммниками AI/AO: ");
            if (aiao.Status != PromptStatus.OK)
            {
                return;
            }
            
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Line line = (Line)acTrans.GetObject(aiao.ObjectId, OpenMode.ForRead);
                    Point3d point = line.StartPoint.Y > line.EndPoint.Y ? line.StartPoint : line.EndPoint;
                    //for (int i =0; i<aiCount; i++)
                    {
                        insertModule(acTrans, acModSpace, acDb, point);

                    }
                    acTrans.Commit();
                }
            }
        }

        private static void insertModule(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d point)
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
                        BlockReference br = new BlockReference(point, bt[blockName]);
                        modSpace.AppendEntity(br);
                        acTrans.AddNewlyCreatedDBObject(br, true);
                        foreach(ObjectId id in br.AttributeCollection)
                        {
                            AttributeReference ar = (AttributeReference)acTrans.GetObject(id, OpenMode.ForWrite);
                            ar.TextString = "1";
                        }
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }
    }
}
