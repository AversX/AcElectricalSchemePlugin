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
            foreach (Cable cable in cables)
            {
                editor.WriteMessage(string.Format("\nMark:{0}; Cable:{1}\n", cable.Mark.TextString, cable.TailsNum));
            }
        }
        private static ObjectId getClosestLine(DBText Text)
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_N,КИА_PE,КИА_КАБЕЛЬ_ОДНОЖИЛЬНЫЙ,КИА_КАБЕЛЬ,КИА_МНОГОЖИЛЬНЫЙ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point = Text.Rotation == 0 ?
                        new Point3d(Text.Position.X, Text.Position.Y - 2, Text.Position.Z) : new Point3d(Text.Position.X + 2, Text.Position.Y, Text.Position.Z);
            Point3dCollection points = new Point3dCollection(new Point3d[] { Text.Position, point });
            PromptSelectionResult selRes = editor.SelectFence(points, filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка поиска!\n");
                return ObjectId.Null;
            }
            else if (selRes.Value.GetObjectIds().Count() > 1)
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
            Cable cable;
            cables.Sort(delegate(Cable x, Cable y)
            {
                return y.Mark.Position.X.CompareTo(x.Mark.Position.X);
            });
            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < cables.Count - 1; i++)
                    for (int j = i + 1; j < cables.Count; j++)
                        if (cables[i].Mark.TextString.Equals(cables[j].Mark.TextString))
                        {
                            if (cables[i].CableLine.Id.Equals(cables[j].CableLine.Id))
                            {
                                DBText mark = (DBText)acTrans.GetObject(cables[j].Mark.Id, OpenMode.ForWrite);
                                mark.Color = ((LayerTableRecord)acTrans.GetObject(mark.LayerId, OpenMode.ForRead)).Color;
                                cables.Remove(cables[j]);
                            }
                            else
                            {
                                List<ObjectId> ids = getIntersectWithTerminal(cables[j].CableLine, acTrans);
                                if (ids != null)
                                {
                                    for (int x = 0; x < ids.Count; x++)
                                    {
                                        Circle circle = (Circle)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                        circle.Color = Color.FromRgb(0, 0, 0);
                                        cable = cables[i];
                                        cable.TailsNum++;
                                        if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                            cables[i] = cable;
                                        else
                                            cables[j] = cable;
                                    }
                                    if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                        cables.Remove(cables[j]);
                                    else
                                    {
                                        cables.Remove(cables[i]);
                                        break;
                                    }
                                }

                                ids = getIntersectWithSplitter(cables[j].CableLine, acTrans);
                                if (ids != null)
                                {
                                    for (int x = 0; x < ids.Count; x++)
                                    {
                                        Hatch hatch = (Hatch)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                        hatch.Color = Color.FromRgb(0, 0, 0);
                                        cable = cables[i];
                                        cable.TailsNum++;
                                        if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                            cables[i] = cable;
                                        else
                                            cables[j] = cable;
                                    }
                                    if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                        cables.Remove(cables[j]);
                                    else
                                    {
                                        cables.Remove(cables[i]);
                                        break;
                                    }
                                }
                            }
                        }
                for (int i = 0; i < cables.Count; i++)
                {
                    if (cables[i].TailsNum == 2)
                    {
                        List<ObjectId> ids = getIntersectWithTerminal(cables[i].CableLine, acTrans);
                        if (ids != null)
                            for (int x = 0; x < ids.Count; x++)
                            {
                                Circle circle = (Circle)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                circle.Color = Color.FromRgb(0, 0, 0);
                                cable = cables[i];
                                cable.TailsNum++;
                                cables[i] = cable;
                            }
                        ids = getIntersectWithSplitter(cables[i].CableLine, acTrans);
                        if (ids != null)
                        {
                            for (int x = 0; x < ids.Count; x++)
                            {
                                Hatch hatch = (Hatch)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                hatch.Color = Color.FromRgb(0, 0, 0);
                                cable = cables[i];
                                cable.TailsNum++;
                                cables[i] = cable;
                            }
                        }
                    }
                }
                acTrans.Commit();
                acTrans.Dispose();
            }
            return cables;
        }

        private static List<ObjectId> getIntersectWithTerminal(Line line, Transaction acTrans)
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "CIRCLE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_КЛЕММЫ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point1;
            Point3d point2;
            if (line.Angle == 0 || line.Angle == 180 || line.Angle == 360)
            {
                if (line.StartPoint.X > line.StartPoint.X)
                {
                    point1 = line.StartPoint.Add(new Vector3d(1, 0, 0));
                    point2 = line.EndPoint.Add(new Vector3d(-1, 0, 0));
                }
                else
                {
                    point1 = line.StartPoint.Add(new Vector3d(-1, 0, 0));
                    point2 = line.EndPoint.Add(new Vector3d(1, 0, 0));
                }
            }
            else
            {
                if (line.StartPoint.Y > line.StartPoint.Y)
                {
                    point1 = line.StartPoint.Add(new Vector3d(0, 1, 0));
                    point2 = line.EndPoint.Add(new Vector3d(0, -1, 0));
                }
                else
                {
                    point1 = line.StartPoint.Add(new Vector3d(0, -1, 0));
                    point2 = line.EndPoint.Add(new Vector3d(0, 1, 0));
                }
            }
            Point3dCollection points = new Point3dCollection(new Point3d[] { point1, point2 });
            PromptSelectionResult circleSelRes = editor.SelectFence(points, filter);
            if (circleSelRes.Status != PromptStatus.OK) return null;
            else
            {
                List<ObjectId> circles = circleSelRes.Value.GetObjectIds().ToList();
                for (int i = 0; i < circles.Count; i++)
                {
                    Circle circle1 = (Circle)acTrans.GetObject(circles[i], OpenMode.ForRead);
                    for (int j = i + 1; j < circles.Count; )
                    {
                        Circle circle2 = (Circle)acTrans.GetObject(circles[j], OpenMode.ForRead);
                        if (circle1.Center.Equals(circle2.Center)) circles.RemoveAt(j);
                        else j++;
                    }
                }
                filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
                filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_КЛЕММЫ");
                filter = new SelectionFilter(filterlist);
                for (int i = 0; i < circles.Count; )
                {
                    Circle circle = (Circle)acTrans.GetObject(circles[i], OpenMode.ForRead);
                    point1 = new Point3d(circle.Center.X - 1, circle.Center.Y - 1, 0);
                    point2 = new Point3d(circle.Center.X + 1, circle.Center.Y + 1, 0);
                    points = new Point3dCollection(new Point3d[] { point1, point2 });
                    PromptSelectionResult lineSelRes = editor.SelectCrossingWindow(point1, point2, filter);
                    if (lineSelRes.Status == PromptStatus.OK)
                    {
                        List<ObjectId> lines = lineSelRes.Value.GetObjectIds().ToList();
                        for (int j = 0; j < lines.Count; j++)
                        {
                            Line ln = (Line)acTrans.GetObject(lines[j], OpenMode.ForWrite);
                            if (ln.Angle >= 0.5 && ln.Angle <= 0.8)
                            {
                                ln.Color = Color.FromRgb(255, 255, 255);
                                i++;
                            }
                            else circles.Remove(circles[i]);
                        }
                    }
                    else circles.Remove(circles[i]);
                }
                if (circles.Count > 0) return circles;
                else return null;
            }
        }

        private static List<ObjectId> getIntersectWithSplitter(Line line, Transaction acTrans)
        {
            TypedValue[] filterlist = new TypedValue[2];
            string layerName = ((LayerTableRecord)acTrans.GetObject(line.LayerId, OpenMode.ForRead)).Name;
            filterlist[0] = new TypedValue((int)DxfCode.Start, "HATCH");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, layerName);
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point1;
            Point3d point2;
            if (line.Angle == 0 || line.Angle == 180 || line.Angle == 360)
            {
                if (line.StartPoint.X > line.StartPoint.X)
                {
                    point1 = line.StartPoint.Add(new Vector3d(1, 0, 0));
                    point2 = line.EndPoint.Add(new Vector3d(-1, 0, 0));
                }
                else
                {
                    point1 = line.StartPoint.Add(new Vector3d(-1, 0, 0));
                    point2 = line.EndPoint.Add(new Vector3d(1, 0, 0));
                }
            }
            else
            {
                if (line.StartPoint.Y > line.StartPoint.Y)
                {
                    point1 = line.StartPoint.Add(new Vector3d(0, 1, 0));
                    point2 = line.EndPoint.Add(new Vector3d(0, -1, 0));
                }
                else
                {
                    point1 = line.StartPoint.Add(new Vector3d(0, -1, 0));
                    point2 = line.EndPoint.Add(new Vector3d(0, 1, 0));
                }
            }
            Point3dCollection points = new Point3dCollection(new Point3d[] { point1, point2 });
            PromptSelectionResult hatchleSelRes = editor.SelectFence(points, filter);
            if (hatchleSelRes.Status != PromptStatus.OK) return null;
            else
            {
                List<ObjectId> hatches = hatchleSelRes.Value.GetObjectIds().ToList();
                if (hatches.Count > 0) return hatches;
                else return null;
            }
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
