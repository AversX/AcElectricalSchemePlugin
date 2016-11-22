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

            PromptEntityResult aiao = acDoc.Editor.GetEntity("Выберите текстовое поле содержащее название клеммника AI/AO");
            if (aiao.Status != PromptStatus.OK)
            {
                return;
            }
            PromptEntityResult dido = acDoc.Editor.GetEntity("Выберите текстовое поле содержащее название клеммника DI/DO");
            if (dido.Status != PromptStatus.OK)
            {
                return;
            }

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    List<DBText> aiaoLinks = new List<DBText>();
                    DBText AIAO = (DBText)acTrans.GetObject(aiao.ObjectId, OpenMode.ForRead);
                    TypedValue[] filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Связи");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    Point3d point = AIAO.Position.Add(new Vector3d(-100, -100, 0));
                    PromptSelectionResult selRes = editor.SelectWindow(AIAO.Position, point, filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nОшибка поиска!\n");
                    }
                    else for (int i = 0; i < selRes.Value.Count; i++)
                        {
                            DBText text = (DBText)acTrans.GetObject(selRes.Value[i].ObjectId, OpenMode.ForWrite);
                            aiaoLinks.Add(text);
                        }

                    List<DBText> didoLinks = new List<DBText>();
                    DBText DIDO = (DBText)acTrans.GetObject(dido.ObjectId, OpenMode.ForRead);
                    filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "Связи");
                    filter = new SelectionFilter(filterlist);
                    point = DIDO.Position.Add(new Vector3d(-100, -100, 0));
                    selRes = editor.SelectWindow(DIDO.Position, point, filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nОшибка поиска!\n");
                    }
                    else for (int i = 0; i < selRes.Value.Count; i++)
                        {
                            DBText text = (DBText)acTrans.GetObject(selRes.Value[i].ObjectId, OpenMode.ForWrite);
                            didoLinks.Add(text);
                        }

                    aiaoLinks = aiaoLinks.OrderByDescending(x => x.Position.Y).ThenByDescending(x => x.Position.X).ToList();
                    didoLinks = didoLinks.OrderByDescending(x => x.Position.Y).ThenByDescending(x => x.Position.X).ToList();

                    int k= 7;
                    for(int i=0; i<aiaoLinks.Count-1;)
                    {
                        if (aiaoLinks[i].TextString.Split('/')[1]==aiaoLinks[i+1].TextString.Split('/')[1])
                        {

                        }
                    }
                }
            }
        }
    }
}
