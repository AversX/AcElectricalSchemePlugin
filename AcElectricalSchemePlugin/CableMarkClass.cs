using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
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
        private static List<BlockCouple> blocks;
        private static List<DBText> marks;
        private static List<Cable> cables;
        private static List<Cable> CalculatedCables;
        static public void CalculateMarks()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            acDb = acDoc.Database;
            
            marks = getAllMarks();
            cables = new List<Cable>();
            CalculatedCables = new List<Cable>();
            blocks = getBlocks();

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                acDoc.LockDocument();
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    foreach (DBText mark in marks)
                    {
                        ObjectId id = getClosestLine(mark);
                        if (id != ObjectId.Null)
                        {
                            Line closestLine = (Line)acTrans.GetObject(id, OpenMode.ForWrite);
                            closestLine.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                            DBText Text = (DBText)acTrans.GetObject(mark.Id, OpenMode.ForWrite);
                            Text.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0);
                            cables.Add(new Cable(mark, closestLine));
                        }
                    }
                    acTrans.Commit();
                    acTrans.Dispose();
                }
            }
            //marks = marks.GroupBy(x => x.TextString).Select(group => group.First()).OrderBy(x => x.TextString, new NaturalStringComparer()).ToList();
            cables = SmartDistinct();
            cables.Sort(new CableComparer());
            foreach (Cable cable in cables)
            {
                editor.WriteMessage(string.Format("\nMark:{0}; Cable:{1}       \n", cable.Mark.TextString, cable.TailsNum));//, inBlock(cable)));
            }
        }

        public static List<string> getData()
        {
            List<string> data = new List<string>();
            for (int i=0; i<cables.Count; i++)
                for (int j=0; j<cables[i].TailsNum;j++)
                    data.Add(cables[i].Mark.TextString);
            List<Cable> cable = calculateBlock();
            for (int i=0; i<cables.Count; i++)
                for (int j=0; j<cables[i].TailsNum;j++)
                    data.Add(cables[i].Mark.TextString);
            data.Sort(new NaturalStringComparer());
            return data;
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

        private static List<Cable> SmartDistinct()
        {
            cables.Sort(delegate(Cable x, Cable y)
            {
                return y.CableLine.Length.CompareTo(x.CableLine.Length);
            });
            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < cables.Count; i++)
                   if (!inBlock(cables[i]))
                    {
                        for (int j = i + 1; j < cables.Count; j++)
                        {
                           if (!inBlock(cables[j]))
                            {
                                if (cables[i].Mark.TextString.Equals(cables[j].Mark.TextString))
                                {
                                    if (cables[i].CableLine.Id==cables[j].CableLine.Id)
                                    {
                                        if (cables[i].Mark.Position.X < cables[j].Mark.Position.X)
                                        {
                                            DBText mark = (DBText)acTrans.GetObject(cables[j].Mark.Id, OpenMode.ForWrite);
                                            mark.Color = ((LayerTableRecord)acTrans.GetObject(mark.LayerId, OpenMode.ForRead)).Color;
                                            cables.Remove(cables[j]);
                                        }
                                        else
                                        {
                                            DBText mark = (DBText)acTrans.GetObject(cables[j].Mark.Id, OpenMode.ForWrite);
                                            mark.Color = ((LayerTableRecord)acTrans.GetObject(mark.LayerId, OpenMode.ForRead)).Color;
                                            cables.Remove(cables[i]);
                                        }
                                    }
                                    else
                                    {
                                        int removed = 0;
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
                                                   // cables.RemoveAt(j);
                                                    removed++;
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
                                                  //  cables.RemoveAt(i);
                                                    removed++;
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
                                                  //  cables.RemoveAt(j);
                                                    removed++;
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
                                                  //  cables.RemoveAt(i);
                                                    removed++;
                                                }
                                            }
                                        }
                                        //if (removed == 0) j++;
                                    }
                                } //else j++;
                            } else cables.RemoveAt(j);
                        } //i++;
                    } else cables.RemoveAt(i);
                for (int i = 0; i < cables.Count; i++)
                {
                    if (cables[i].TailsNum == 2)
                        calculateCable(i, acTrans);
                }
                acTrans.Commit();
                acTrans.Dispose();
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

        private static List<Cable> calculateBlock()
        {
            List<Cable> calculated = new List<Cable>();
            for (int i = 0; i < blocks.Count; i++)
            {
                for (int j = 0; j < blocks[i].Cables.Count; j++)
                {
                    Cable cableL = blocks[i].Cables[j].cable;
                    for (int k = j + 1; k < blocks[i].Cables.Count; k++)
                    {
                        if (!blocks[i].Cables[k].Left)
                        {
                            Cable cableR = blocks[i].Cables[k].cable;
                            if (blocks[i].Cables[j].Left)
                            {
                                if (cableL.Mark.Position.Y >= cableR.Mark.Position.Y - 1 && cableL.Mark.Position.Y <= cableR.Mark.Position.Y + 1)
                                    if (cableL.Mark.TextString[0] == cableR.Mark.TextString[0] && cableL.Mark.TextString[1] == cableR.Mark.TextString[1])
                                    {
                                        string left = cableL.Mark.TextString;
                                        string right = cableR.Mark.TextString;
                                        left.Remove(0, 2);
                                        right.Remove(0, 2);
                                        int leftCount;
                                        int rightCount;
                                        int.TryParse(left, out leftCount);
                                        int.TryParse(right, out rightCount);
                                        if (leftCount != 0 && rightCount != 0)
                                        {
                                            if (leftCount < rightCount)
                                                for (int q = leftCount; q <= rightCount; q++)
                                                {
                                                    DBText mark = new DBText();
                                                    mark.TextString = cableL.Mark.TextString[0] + cableL.Mark.TextString[1] + leftCount.ToString();
                                                    Cable cable = new Cable(mark, new Line());
                                                    cable.TailsNum = cableL.TailsNum;
                                                    if (!calculated.Exists(x => x.Mark.TextString == mark.TextString))
                                                        calculated.Add(cable);
                                                }
                                        }
                                    }
                                    else if (cableL.Mark.TextString == cableR.Mark.TextString)
                                        for (int q = 0; q < blocks[i].Number; q++)
                                        {
                                            DBText mark = new DBText();
                                            mark.TextString = cableL.Mark.TextString;
                                            Cable cable = new Cable(mark, new Line());
                                            cable.TailsNum = cableL.TailsNum;
                                            if (!calculated.Exists(x => x.Mark.TextString == mark.TextString))
                                                calculated.Add(cable);
                                        }
                            }
                            else
                            {
                                if (cableL.Mark.Position.Y >= cableR.Mark.Position.Y - 1 && cableL.Mark.Position.Y <= cableR.Mark.Position.Y + 1)
                                    if (cableL.Mark.TextString[0] == cableR.Mark.TextString[0] && cableL.Mark.TextString[1] == cableR.Mark.TextString[1])
                                    {
                                        string left = cableR.Mark.TextString;
                                        string right = cableL.Mark.TextString;
                                        left.Remove(0, 2);
                                        right.Remove(0, 2);
                                        int leftCount;
                                        int rightCount;
                                        int.TryParse(left, out leftCount);
                                        int.TryParse(right, out rightCount);
                                        if (leftCount != 0 && rightCount != 0)
                                        {
                                            if (leftCount < rightCount)
                                                for (int q = leftCount; q <= rightCount; q++)
                                                {
                                                    DBText mark = new DBText();
                                                    mark.TextString = cableL.Mark.TextString[0] + cableL.Mark.TextString[1] + leftCount.ToString();
                                                    Cable cable = new Cable(mark, new Line());
                                                    cable.TailsNum = cableL.TailsNum;
                                                    if (!calculated.Exists(x => x.Mark.TextString == mark.TextString))
                                                        calculated.Add(cable);
                                                }
                                        }
                                    }
                                    else if (cableL.Mark.TextString == cableR.Mark.TextString)
                                        for (int q = 0; q < blocks[i].Number; q++)
                                        {
                                            DBText mark = new DBText();
                                            mark.TextString = cableL.Mark.TextString;
                                            Cable cable = new Cable(mark, new Line());
                                            cable.TailsNum = cableL.TailsNum;
                                            if (!calculated.Exists(x => x.Mark.TextString == mark.TextString))
                                                calculated.Add(cable);
                                        }
                            }
                        }
                    }
                }
            }
            return calculated;
        }
        
        private static Cable findCableByMark(List<Cable> cables, DBText mark)
        {
            Cable result = new Cable();
            for (int i = 0; i < cables.Count; i++)
                if (mark.Equals(cables[i].Mark)) result = cables[i];
            return result;
        }

        private static bool inBlock(Cable cable)
        {
            BlockCouple block;
            for (int i = 0; i < blocks.Count; i++)
            {
                if (inPerimeter(blocks[i].Block1, cable.Mark.Position))
                {
                    block = blocks[i];
                    List<BlockCable> bC = block.Cables;
                    bC.Add(new BlockCable(cable, true));
                    block.Cables = bC;
                    return true;
                }
                else if (inPerimeter(blocks[i].Block2, cable.Mark.Position))
                {
                    block = blocks[i];
                    List<BlockCable> bC = block.Cables;
                    bC.Add(new BlockCable(cable, false));
                    block.Cables = bC;
                    return true;
                }
            }
            return false;
        }

        private static bool inPerimeter(Polyline poly, Point3d point)
        {
            Point3d point1 = poly.GetPoint3dAt(0);
            for (int i = 1; i < poly.NumberOfVertices; i++)
            {
                Point3d p = poly.GetPoint3dAt(i);
                if (point1.X >= p.X && point1.Y <= p.Y)
                    point1 = p;
            }
            Point3d point3 = poly.GetPoint3dAt(0);
            for (int i = 1; i < poly.NumberOfVertices; i++)
            {
                Point3d p = poly.GetPoint3dAt(i);
                if (point3.X <= p.X && point3.Y >= p.Y)
                    point3 = p;
            }
            if (point.X >= point1.X && point.X <= point3.X && point.Y <= point1.Y && point.Y >= point3.Y) 
                return true;
            else return false;
        }

        private static List<BlockCouple> getBlocks()
        {
            List<BlockCouple> blocks = new List<BlockCouple>(); 
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ПУНКТИР");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectAll(filter);
            List<Polyline> polys = new List<Polyline>();
            if (selRes.Status == PromptStatus.OK)
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    List<ObjectId> objIds = selRes.Value.GetObjectIds().ToList();
                    for (int i = 0; i < objIds.Count; i++)
                    {
                        Polyline poly = (Polyline)acTrans.GetObject(objIds[i], OpenMode.ForRead);
                        if (poly.Area > 30000) polys.Add(poly);
                    }
                    acTrans.Commit();
                }
            }
            for (int i = 0; i < polys.Count; i++)
            {
                for (int j = i + 1; j < polys.Count; j++)
                {
                    if (polys[i].StartPoint.Y <= polys[j].StartPoint.Y + 1 && polys[i].StartPoint.Y >= polys[j].StartPoint.Y - 1)
                    {
                        if (comparePolys(polys[i], polys[j]))
                        {
                            blocks.Add(new BlockCouple(polys[i], polys[j]));
                            BlockCouple bc = blocks[blocks.Count - 1];
                            bc.Number = getNumOfBlocks(bc);
                            blocks[blocks.Count - 1] = bc;
                        }
                        else if (comparePolys(polys[j], polys[i]))
                        {
                            blocks.Add(new BlockCouple(polys[j], polys[i]));
                            BlockCouple bc = blocks[blocks.Count - 1];
                            bc.Number = getNumOfBlocks(bc);
                            blocks[blocks.Count - 1] = bc;
                        }
                        else j++;
                    }
                    else j++;
                }
            }
            return blocks;
        }

        private static int getNumOfBlocks(BlockCouple blockCouple)
        {
            int result = 0;
            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ПУНКТИР");
            SelectionFilter filter = new SelectionFilter(filterlist);
            Point3d point1 = blockCouple.Block1.GetPoint3dAt(0);
            for (int i = 1; i < blockCouple.Block1.NumberOfVertices; i++)
            {
                Point3d point = blockCouple.Block1.GetPoint3dAt(i);
                if (point.X >= point1.X && point.Y >= point1.Y)
                    point1 = point;
            }
            Point3d point2 = blockCouple.Block2.GetPoint3dAt(0);
            for (int i = 1; i < blockCouple.Block2.NumberOfVertices; i++)
            {
                Point3d point = blockCouple.Block2.GetPoint3dAt(i);
                if (point.X <= point2.X && point.Y <= point2.Y)
                    point2 = point;
            }
            PromptSelectionResult selRes = editor.SelectCrossingWindow(point1, point2, filter);
            if (selRes.Status == PromptStatus.OK)
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    Line line = (Line)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForRead);
                    filterlist = new TypedValue[1];
                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                    filter = new SelectionFilter(filterlist);
                    point1 = line.StartPoint.Add(new Vector3d(0, 10, 0));
                    point2 = line.EndPoint.Add(new Vector3d(0, -10, 0));
                    selRes = editor.SelectCrossingWindow(point1, point2);
                    if (selRes.Status == PromptStatus.OK)
                    {
                        DBText dbText = (DBText)acTrans.GetObject(selRes.Value.GetObjectIds()[0], OpenMode.ForRead);
                        string text = dbText.TextString;
                        int.TryParse(text, out result);
                    }
                    acTrans.Commit();
                }
            }
            return result;
        }

        private static bool comparePolys(Polyline poly1, Polyline poly2)
        {
            double X1 = poly1.StartPoint.X;
            double X2 = poly2.StartPoint.X;
            if (X1 < X2)
            {
                double delta = Math.Abs(X2) - Math.Abs(X1);
                if (delta <= 1000 && delta>0)
                    return true;
                else return false;
            }
            else return false;
        }
    }

    class CableComparer : IComparer<Cable>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);
        public int Compare(Cable x, Cable y)
        {
            return StrCmpLogicalW(x.Mark.TextString, y.Mark.TextString);
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
    public struct BlockCable
    {
        public Cable cable;
        public bool Left;
        public BlockCable(Cable _cable, bool left)
        {
            cable = _cable;
            Left = left;
        }

    }

    public struct BlockCouple
    {
        public Polyline Block1;
        public Polyline Block2;
        public int Number;
        public List<BlockCable> Cables;
        public BlockCouple(Polyline block1, Polyline block2)
        {
            Block1 = block1;
            Block2 = block2;
            Number = 0;
            Cables = new List<BlockCable>();
        }
    }
}
