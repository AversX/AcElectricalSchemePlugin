using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AcElectricalSchemePlugin
{
    static class CableMark
    {
        private static Editor editor;
        private static Database acDb;
        private static List<DBText> marks;
        private static List<DBText> dbtexts;
        private static List<MText> mtexts;
        private static List<Cable> cables;

        static public void CalculateMarks()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            acDb = acDoc.Database;

            cables = new List<Cable>();
            
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    marks = getAllMarks(acTrans);
                    if(marks!=null)
                        foreach (DBText mark in marks)
                        {
                            ObjectId id = getClosestLine(mark);
                            if (id != ObjectId.Null)
                            {
                                Line closestLine = (Line)acTrans.GetObject(id, OpenMode.ForWrite);
                                closestLine.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                DBText Text = (DBText)acTrans.GetObject(mark.Id, OpenMode.ForWrite);
                                Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                cables.Add(new Cable(mark.TextString, mark.Id, mark.Position, closestLine));
                            }
                        }
                    cables = SmartDistinct(acTrans);
                    dbtexts = getAllDBTexts(acTrans);
                    if (dbtexts != null)
                        foreach (DBText text in dbtexts)
                        {
                            if (text.TextString.Length>0)
                            {
                                if (text.TextString[0] == '-' || text.TextString.Contains("WA"))
                                {
                                    DBText Text = (DBText)acTrans.GetObject(text.Id, OpenMode.ForWrite);
                                    Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                    cables.Add(new Cable(text.TextString, text.Id, text.Position, null));
                                }
                            }
                        }
                    mtexts = getAllMTexts(acTrans);
                    if (mtexts != null)
                        foreach (MText text in mtexts)
                        {
                            if (text.Text.Length > 0)
                            {
                                if (text.Text[0] == '-' || text.Text.Contains("WA"))
                                {
                                    MText Text = (MText)acTrans.GetObject(text.Id, OpenMode.ForWrite);
                                    Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                                    cables.Add(new Cable(text.Text, text.Id, text.Location, null));
                                }
                            }
                        }
                    cables.Sort(new CableComparer());
                    acTrans.Commit();
                    acTrans.Dispose();
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text|*.txt";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
                System.IO.File.WriteAllLines(saveFileDialog1.FileName, getData());
        }

        private static List<string> getData()
        {
            List<string> data = new List<string>();
            for (int i=0; i<cables.Count; i++)
                for (int j=0; j<cables[i].TailsNum;j++)
                    data.Add(cables[i].Mark);
            data.Sort(new NaturalStringComparer());
            return data;
        }

        private static ObjectId getClosestLine(DBText Text)
        {
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_N,КИА_PE,КИА_КАБЕЛЬ_ОДНОЖИЛЬНЫЙ,КИА_КАБЕЛЬ,КИА_провод,КИА_МНОГОЖИЛЬНЫЙ,КИА_КАБЕЛЬ_24В,КИА_КАБЕЛБ_24В,0");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point;
            if (Text.Rotation == 0)
            {
                point = new Point3d(Text.Position.X + 3, Text.Position.Y - 2, Text.Position.Z);
            }
            else
            {
                point = new Point3d(Text.Position.X + 2, Text.Position.Y + 3, Text.Position.Z);
            }
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

        private static List<DBText> getAllMarks(Transaction acTrans)
        {
            List<DBText> marks = new List<DBText>();
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_МАРКИРОВКА");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка выборки маркировок!\n");
                return null;
            }
            foreach (ObjectId id in selRes.Value.GetObjectIds())
            {
                marks.Add((DBText)acTrans.GetObject(id, OpenMode.ForRead));
            }
            return marks;
        }

        private static List<DBText> getAllDBTexts(Transaction acTrans)
        {
            List<DBText> texts = new List<DBText>();
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ТЕКСТ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка выборки текстовых полей!\n");
                return null;
            }
            foreach (ObjectId id in selRes.Value.GetObjectIds())
            {
                texts.Add((DBText)acTrans.GetObject(id, OpenMode.ForWrite));
            }
            return texts;
        }

        private static List<MText> getAllMTexts(Transaction acTrans)
        {
            List<MText> texts = new List<MText>();
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ТЕКСТ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка выборки текстовых полей!\n");
                return null;
            }
            foreach (ObjectId id in selRes.Value.GetObjectIds())
            {
                texts.Add((MText)acTrans.GetObject(id, OpenMode.ForWrite));
            }
            return texts;
        }

        private static List<Cable> SmartDistinct(Transaction acTrans)
        {
            cables.Sort(delegate(Cable x, Cable y)
            {
                return y.CableLine.Length.CompareTo(x.CableLine.Length);
            });
            for (int i = 0; i < cables.Count; i++)
                for (int j = i + 1; j < cables.Count; j++)
                {
                    if (cables[i].Mark.Equals(cables[j].Mark))
                    {
                        if (cables[i].CableLine.Id == cables[j].CableLine.Id)
                        {
                            DBText mark = (DBText)acTrans.GetObject(cables[j].Id, OpenMode.ForWrite);
                            mark.Color = Color.FromRgb(255,255,255);
                            cables.Remove(cables[j]);
                            j--;
                        }
                        else
                        {
                            List<ObjectId> ids1 = getIntersectWithTerminal(cables[i].CableLine, acTrans);
                            List<ObjectId> ids2 = getIntersectWithTerminal(cables[j].CableLine, acTrans);
                            if (ids1 != null && ids2 != null)
                            {
                                List<ObjectId> ids = ids1.Intersect(ids2).ToList();
                                if (ids != null)
                                {
                                    if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                    {
                                        ids.AddRange(ids1.Except(ids2));
                                        for (int x = 0; x < ids.Count; x++)
                                        {
                                            Circle circle = (Circle)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                            circle.Color = Color.FromRgb(0, 0, 0);
                                            Cable cable;
                                            cable = cables[i];
                                            cable.TailsNum++;
                                            cables[i] = cable;
                                        }
                                    }
                                    else
                                    {
                                        ids.AddRange(ids2.Except(ids1));
                                        for (int x = 0; x < ids.Count; x++)
                                        {
                                            Circle circle = (Circle)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                            circle.Color = Color.FromRgb(0, 0, 0);
                                            Cable cable;
                                            cable = cables[j];
                                            cable.TailsNum++;
                                            cables[j] = cable;
                                        }
                                    }
                                }
                            }
                            ids1 = getIntersectWithSplitter(cables[i].CableLine, acTrans);
                            ids2 = getIntersectWithSplitter(cables[j].CableLine, acTrans);
                            if (ids1 != null && ids2 != null)
                            {
                                List<ObjectId> ids = ids1.Intersect(ids2).ToList();
                                if (ids != null)
                                {
                                    if (cables[i].CableLine.Length > cables[j].CableLine.Length)
                                    {
                                        ids.AddRange(ids1.Except(ids2));
                                        for (int x = 0; x < ids.Count; x++)
                                        {
                                            Hatch hatch = (Hatch)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                            hatch.Color = Color.FromRgb(0, 0, 0);
                                            Cable cable;
                                            cable = cables[i];
                                            cable.TailsNum++;
                                            cables[i] = cable;
                                        }
                                    }
                                    else
                                    {
                                        ids.AddRange(ids2.Except(ids1));
                                        for (int x = 0; x < ids.Count; x++)
                                        {
                                            Hatch hatch = (Hatch)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                                            hatch.Color = Color.FromRgb(0, 0, 0);
                                            Cable cable;
                                            cable = cables[j];
                                            cable.TailsNum++;
                                            cables[j] = cable;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            for (int i = 0; i < cables.Count; i++)
            {
                if (cables[i].TailsNum == 2)
                    calculateCable(i, acTrans);
            }
            return cables;
        }

        private static void calculateCable(int indexI, Transaction acTrans)
        {
            List<ObjectId> ids = getIntersectWithTerminal(cables[indexI].CableLine, acTrans);
            if (ids != null)
            {
                for (int x = 0; x < ids.Count; x++)
                {
                    Circle circle = (Circle)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                    circle.Color = Color.FromRgb(0, 0, 0);
                    Cable cable;
                    cable = cables[indexI];
                    cable.TailsNum++;
                    cables[indexI] = cable;
                }
            }
            ids = getIntersectWithSplitter(cables[indexI].CableLine, acTrans);
            if (ids != null)
            {
                for (int x = 0; x < ids.Count; x++)
                {
                    Hatch hatch = (Hatch)acTrans.GetObject(ids[x], OpenMode.ForWrite);
                    hatch.Color = Color.FromRgb(0, 0, 0);
                    Cable cable;
                    cable = cables[indexI];
                    cable.TailsNum++;
                    cables[indexI] = cable;
                }
            }
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
                if (line.StartPoint.X > line.EndPoint.X)
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
                if (line.StartPoint.Y > line.EndPoint.Y)
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
                    Circle circle1 = (Circle)acTrans.GetObject(circles[i], OpenMode.ForWrite);
                    for (int j = i + 1; j < circles.Count; j++)
                    {
                        Circle circle2 = (Circle)acTrans.GetObject(circles[j], OpenMode.ForWrite);
                        if (circle1.Center.Equals(circle2.Center)) { circles.RemoveAt(j); j--; }
                    }
                }
                filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
                filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_КЛЕММЫ");
                filter = new SelectionFilter(filterlist);
                for (int i = 0; i < circles.Count; i++)
                {
                    Circle circle = (Circle)acTrans.GetObject(circles[i], OpenMode.ForWrite);
                    point1 = new Point3d(circle.Center.X - 2, circle.Center.Y - 2, 0);
                    point2 = new Point3d(circle.Center.X + 2, circle.Center.Y + 2, 0);
                    points = new Point3dCollection(new Point3d[] { point1, point2 });
                    PromptSelectionResult lineSelRes = editor.SelectCrossingWindow(point1, point2, filter);
                    if (lineSelRes.Status == PromptStatus.OK)
                    {
                        List<ObjectId> lines = lineSelRes.Value.GetObjectIds().ToList();
                        bool terminal = false;
                        for (int j = 0; j < lines.Count; j++)
                        {
                            Line ln = (Line)acTrans.GetObject(lines[j], OpenMode.ForWrite);
                            if (ln.Angle >= 0.5 && ln.Angle <= 0.8)
                            {
                                ln.Color = Color.FromRgb(255, 255, 255);
                                circle.Color = Color.FromRgb(255, 255, 255);
                                terminal = true;
                                break;
                            }
                            else terminal = false;
                        }
                        if (terminal) { circles.Remove(circles[i]); i--; }
                    }
                    else { circles.Remove(circles[i]); i--; }
                }
                if (circles.Count > 0) return circles;
                else return null;
            }
        }

        private static List<ObjectId> getIntersectWithSplitter(Line line, Transaction acTrans)
        {
            TypedValue[] filterlist = new TypedValue[2];
            string layerName = ((LayerTableRecord)acTrans.GetObject(line.LayerId, OpenMode.ForWrite)).Name;
            filterlist[0] = new TypedValue((int)DxfCode.Start, "HATCH");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point1;
            Point3d point2;
            if (line.Angle == 0 || line.Angle == 180 || line.Angle == 360)
            {
                if (line.StartPoint.X > line.EndPoint.X)
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
                if (line.StartPoint.Y > line.EndPoint.Y)
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

    class CableComparer : IComparer<Cable>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);
        public int Compare(Cable x, Cable y)
        {
            return StrCmpLogicalW(x.Mark, y.Mark);
        }
    }

    class NaturalStringComparer : IComparer<String>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);
        public int Compare(string x, string y)
        {
            return StrCmpLogicalW(x, y);
        }
    }

    public struct Cable
    {
        public string Mark;
        public ObjectId Id;
        public Point3d Position;
        public int TailsNum;
        public Line CableLine;
        public Cable(string mark, ObjectId id, Point3d position, Line cableLine)
        {
            Mark = mark;
            Id = id;
            Position = position;
            TailsNum = 2;
            CableLine = cableLine;
        }
    }
}
