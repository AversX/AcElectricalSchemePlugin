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
        static Editor editor;
        static Document acDoc;

        private struct Cable
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

        static public void CalculateMarks()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;
            List<ObjectId> lines = getAllLines();
            List<DBText> marks = getAllMarks(acDb);
            marks = marks.GroupBy(x => x.TextString).Select(group => group.First()).OrderBy(x => x.TextString, new NaturalStringComparer()).ToList();
            List<Cable> cables = new List<Cable>();
            foreach (DBText text in marks)
            {
                Line closestLine = getClosestLine(lines, text, acDb);
                if (closestLine != null)
                    try
                    {
                        using (DocumentLock docLock = acDoc.LockDocument())
                        {
                            acDoc.LockDocument();
                            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                            {
                                cables.Add(new Cable(text, closestLine));
                                
                                Line line = (Line)acTrans.GetObject(closestLine.Id, OpenMode.ForWrite);
                                closestLine.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                DBText Text = (DBText)acTrans.GetObject(text.Id, OpenMode.ForWrite);
                                Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);

                                lines.Remove(closestLine.Id);

                                acTrans.Commit();
                                acTrans.Dispose();
                            }
                        }
                    }
                    catch
                    {
                        editor.WriteMessage("Ошибка при попытке перекрасить");
                    }
            }
            foreach(Cable cable in cables)
            {
                editor.WriteMessage(string.Format("\nMark:{0}\n", cable.Mark.TextString));
            }
        }
        private static Line getClosestLine(List<ObjectId> Lines, DBText Text, Database AcDb)
        {
            using (Transaction acTrans = AcDb.TransactionManager.StartTransaction())
            {
                try
                {
                    Line closestLine = null;
                    foreach (ObjectId objId in Lines)
                    {
                        Line line = (Line)acTrans.GetObject(objId, OpenMode.ForRead);
                        Point3d point = Text.Rotation == 0 ?
                            new Point3d(Text.Position.X, Text.Position.Y - 2, Text.Position.Z) : new Point3d(Text.Position.X + 2, Text.Position.Y, Text.Position.Z);
                        Line cross = new Line(Text.Position, point);
                        Point3dCollection pts = new Point3dCollection();
                        line.IntersectWith(cross, Intersect.OnBothOperands, pts, IntPtrCollection.DefaultSize, IntPtrCollection.DefaultSize);
                        if (pts.Count > 0)
                        {
                            closestLine = line;
                        }
                    }
                    if (closestLine == null)
                    {
                        editor.WriteMessage("Кабель не найден для маркировки {0}", Text.TextString);
                        acTrans.Commit();
                        acTrans.Dispose();
                        return null;
                    }
                    else
                    {
                        acTrans.Commit();
                        acTrans.Dispose();
                        return closestLine;
                    }
                }
                catch
                {
                    editor.WriteMessage("Ошибка поиска кабеля для маркировки {0}", Text.TextString);
                    acTrans.Commit();
                    acTrans.Dispose();
                    return null;
                }
            }
        }

        private static List<ObjectId> getAllLines()
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue(0, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_N,КИА_PE,КИА_КАБЕЛЬ_ОДНОЖИЛЬНЫЙ,КИА_КАБЕЛЬ,КИА_МНОГОЖИЛЬНЫЙ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка выборки линий кабелей!\n");
                return null;
            }
            else return selRes.Value.GetObjectIds().ToList();
        }

        private static List<DBText> getAllMarks(Database acDb)
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
    }
    class NaturalStringComparer : IComparer<string>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);

        public int Compare(string x, string y)
        {
            return StrCmpLogicalW(x, y);
        }
    }
}
