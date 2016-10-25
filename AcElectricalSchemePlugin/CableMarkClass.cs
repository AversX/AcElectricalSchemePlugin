using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;

namespace AcElectricalSchemePlugin
{
    static class CableMark
    {
        private static Editor editor;
        private static Document acDoc;
        private static Database acDb;

        static public void CalculateMarks()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            acDb = acDoc.Database;
            List<DBText> marks = getAllMarks();
            List<Cable> cables = new List<Cable>();
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                acDoc.LockDocument();
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    foreach (DBText mark in marks)
                    {
                        ObjectId id = getClosestLine(mark);
                        if (id != ObjectId.Null)
                            try
                            {
                                Line closestLine = (Line)acTrans.GetObject(id, OpenMode.ForWrite);
                                closestLine.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                DBText Text = (DBText)acTrans.GetObject(mark.Id, OpenMode.ForWrite);
                                Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                cables.Add(new Cable(mark, closestLine));
                            }
                            catch
                            {
                                editor.WriteMessage("Ошибка при попытке перекрасить");
                            }
                    }
                    acTrans.Commit();
                    acTrans.Dispose();
                }
            }
            //marks = marks.GroupBy(x => x.TextString).Select(group => group.First()).OrderBy(x => x.TextString, new NaturalStringComparer()).ToList();
            cables = SmartDistinct(cables);
            cables.Sort(new NaturalStringComparer());
            foreach(Cable cable in cables)
            {
                editor.WriteMessage(string.Format("\nMark:{0}; Cable:{1}\n", cable.Mark.TextString, cable.CableLine.Id));
            }
        }
        private static ObjectId getClosestLine(DBText Text)
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue(0, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_N,КИА_PE,КИА_КАБЕЛЬ_ОДНОЖИЛЬНЫЙ,КИА_КАБЕЛЬ,КИА_МНОГОЖИЛЬНЫЙ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point = Text.Rotation == 0 ?
                        new Point3d(Text.Position.X, Text.Position.Y - 2, Text.Position.Z) : new Point3d(Text.Position.X + 2, Text.Position.Y, Text.Position.Z);
            Point3dCollection points = new Point3dCollection(new Point3d[]{Text.Position, point});
            PromptSelectionResult selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка поиска! selRes.Status.ToString()\n");
                return ObjectId.Null;
            }
            else if (selRes.Value.GetObjectIds().Count() >1)
            {
                editor.WriteMessage("\nНайдено больше одного кабеля!\n");
                return ObjectId.Null;
            }
            else if (selRes.Value.GetObjectIds().Count() == 0)
            {
                editor.WriteMessage("\nНе найдено ни одного кабеля!\n");
                return ObjectId.Null;
            }
            else return selRes.Value.GetObjectIds().ToList()[0];
            #region Старый метод поиска кабеля
            //using (Transaction acTrans = AcDb.TransactionManager.StartTransaction())
            //{
            //    try
            //    {
            //        Line closestLine = null;
            //        foreach (ObjectId objId in Lines)
            //        {
            //            Line line = (Line)acTrans.GetObject(objId, OpenMode.ForRead);
            //            Point3d point = Text.Rotation == 0 ?
            //                new Point3d(Text.Position.X, Text.Position.Y - 2, Text.Position.Z) : new Point3d(Text.Position.X + 2, Text.Position.Y, Text.Position.Z);
            //            Line cross = new Line(Text.Position, point);
            //            Point3dCollection pts = new Point3dCollection();
            //            line.IntersectWith(cross, Intersect.OnBothOperands, pts, IntPtrCollection.DefaultSize, IntPtrCollection.DefaultSize);
            //            if (pts.Count > 0)
            //            {
            //                closestLine = line;
            //            }
            //        }
            //        if (closestLine == null)
            //        {
            //            editor.WriteMessage("Кабель не найден для маркировки {0}", Text.TextString);
            //            acTrans.Commit();
            //            acTrans.Dispose();
            //            return null;
            //        }
            //        else
            //        {
            //            acTrans.Commit();
            //            acTrans.Dispose();
            //            return closestLine;
            //        }
            //    }
            //    catch
            //    {
            //        editor.WriteMessage("Ошибка поиска кабеля для маркировки {0}", Text.TextString);
            //        acTrans.Commit();
            //        acTrans.Dispose();
            //        return null;
            //    }
            //}
            //private static List<ObjectId> getAllLines()
            //{
            //    TypedValue[] filterlist = new TypedValue[2];
            //    filterlist[0] = new TypedValue(0, "LINE");
            //    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_N,КИА_PE,КИА_КАБЕЛЬ_ОДНОЖИЛЬНЫЙ,КИА_КАБЕЛЬ,КИА_МНОГОЖИЛЬНЫЙ");
            //    SelectionFilter filter = new SelectionFilter(filterlist);
            //    PromptSelectionResult selRes = editor.SelectAll(filter);
            //    if (selRes.Status != PromptStatus.OK)
            //    {
            //        editor.WriteMessage("\nОшибка выборки линий кабелей!\n");
            //        return null;
            //    }
            //    else return selRes.Value.GetObjectIds().ToList();
            //}
#endregion
        }

        private static List<DBText> getAllMarks()
        {
            List<DBText> marks = new List<DBText>();
            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                TypedValue[] filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_МАРКИРОВКА");
                SelectionFilter filter = new SelectionFilter(filterlist);
                PromptSelectionResult selRes = editor.SelectAll(filter);
                if (selRes.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("\nОшибка выборки маркировок!\n");
                    acTrans.Commit();
                    return null;
                }
                foreach (ObjectId id in selRes.Value.GetObjectIds())
                {
                    marks.Add((DBText)acTrans.GetObject(id, OpenMode.ForRead));
                } 
                acTrans.Commit();
                return marks;
            }
        }

        private static List<Cable> SmartDistinct(List<Cable> Cables)
        {
            List<Cable> cables = Cables;
            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < cables.Count; i++)
                    for (int j = 1; j < cables.Count; j++)
                        if (cables[i].Mark.TextString.Equals(cables[j].Mark.TextString))
                            if (cables[i].CableLine.Id.Equals(cables[j].CableLine.Id))
                            {
                                DBText mark = (DBText)acTrans.GetObject(cables[j].Mark.Id, OpenMode.ForWrite);
                                mark.Color = ((LayerTableRecord)acTrans.GetObject(mark.LayerId, OpenMode.ForRead)).Color;
                                cables.Remove(cables[j]);
                            }
                acTrans.Commit();
                acTrans.Dispose();
            }
            return cables;
        }
    }

    class NaturalStringComparer : IComparer<Cable>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);

        public int Compare(Cable x, Cable y)
        {
            return StrCmpLogicalW(x.Mark.TextString, y.Mark.TextString);
        }
    }

    public struct Cable
    {
        public DBText Mark;
        public int TailsNum;
        public Line CableLine;
        public Cable(DBText mark, Line cableLine)
        {
            Mark = mark;
            TailsNum = 2;
            CableLine = cableLine;
        }
    }
}
