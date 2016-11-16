using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.EditorInput;

namespace AcElectricalSchemePlugin
{
    static class ConnectionScheme
    {
        private static Document acDoc;
        private static List<Unit> units;
        private static List<tBoxUnit> tBoxes;
        private static Editor editor;
        private static Table currentTable = null;
        private static bool firstSheet;
        private static LinetypeTable lineTypeTable;
        private static TextStyleTableRecord textStyle;
        private static Polyline currentTBox = null;
        private static Unit currentUnit;
        private static MText currentTBoxName = null;
        private static Leader currentLeader = null;
        private static double width = 0;
        private static List<Line> cableLineUps;
        private static List<MText> cableLineText;
        private static List<Table> tables = new List<Table>();

        struct tBoxUnit
        {
            public string Name;
            public int Count;
            public int LastTerminalNumber;
            public MText textName;
            public Leader ldr;

            public tBoxUnit(string name, int count, int lastTerminalNumber)
            {
                Name = name;
                Count = count;
                LastTerminalNumber = lastTerminalNumber;
                textName = null;
                ldr = null;
            }
        }

        struct Unit
        {
            public string cupboardName;
            public string tBoxName;
            public string designation;
            public string param;
            public string equipment;
            public string equipType;
            public string cableMark;
            public bool shield;
            public List<string> terminals;
            public List<string> equipTerminals;
            public Line cableMarkLine;
            public Line outputCableLine;
            public MText cableName;
            public bool newSheet; 

            public Unit(string _cupboardName, string _tBoxName, string _designation, string _param, string _equipment, string _equipType, string _cableMark, bool _shield, List<string> _terminals, List<string> _equipTerminals)
            {
                cupboardName = _cupboardName;
                tBoxName = _tBoxName;
                designation = _designation;
                param = _param;
                equipment = _equipment;
                equipType = _equipType;
                cableMark = _cableMark;
                shield = _shield;
                terminals = _terminals;
                equipTerminals = _equipTerminals;
                cableMarkLine = null;
                outputCableLine = null;
                cableName = null;
                newSheet = false;
            }
        }

        private static void findLines(string[] lines, int index)
        {
            if (lines[index].Contains("\""))
            {
                for (int j = index + 1; j < lines.Length; j++)
                {
                    List<int> indexes = new List<int>();
                    for (int c = 0; c < lines[j].Length; c++)
                    {
                        if (lines[j][c] == '\"') indexes.Add(c);
                    }
                    if (indexes.Count == 0) { lines[index] += " "+lines[j]; lines[j] = ""; }
                    else if (indexes.Count == 1)
                    {
                        lines[index] += " " + lines[j];
                        lines[j] = "";
                        break;
                    }
                    else if (indexes.Count > 1)
                    {
                        findLines(lines, j);
                        lines[index] += " " + lines[j]; 
                        lines[j] = "";
                        break;
                    }
                }
            }
        }
        private static List<int> findGroups(int index)
        {
            List<int> result = new List<int>();
            result.Add(index);
            for (int i = index+1; i < units.Count; i++)
            {
                if (units[index].tBoxName == units[i].tBoxName && units[index].tBoxName != "" && units[i].tBoxName != "")
                    result.Add(i);
                else break;
            }
            return result.Count == 1 ? null : result;
        }
        public static void DrawScheme()
        {
            units = new List<Unit>();
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "(*.csv)|*.csv";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] lines = System.IO.File.ReadLines(file.FileName, System.Text.Encoding.GetEncoding(1251)).ToArray();
                if (lines.Length > 0)
                {
                    for (int i = 2; i < lines.Length; i++)
                    {
                        if (lines[i]!="")
                            findLines(lines, i);
                    }
                    for (int i = 2; i < lines.Length; i++)
                    {
                        if (lines[i] != "")
                        {
                            string[] text = lines[i].Split(';');
                            List<string> terminals = new List<string>();
                            for (int j = 8; j < 18; j++)
                                if (text[j] != "") terminals.Add(text[j]);
                            List<string> equipTerminals = new List<string>();
                            for (int j = 18; j < 28; j++)
                                if (text[j] != "") equipTerminals.Add(text[j]);
                            bool shield = text[7].ToUpper() == "ДА" || text[7].ToUpper() == "ЕСТЬ" ? true : false;
                            units.Add(new Unit(text[0], text[1], text[2], text[3], text[4], text[5], text[6], shield, terminals, equipTerminals));
                        }
                    }
                    firstSheet = true;
                    tBoxes = new List<tBoxUnit>();
                    tables = new List<Table>();
                    connectionScheme();
                }
            }
        }
        private static void connectionScheme()
        {
            Database acDb = acDoc.Database;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    PromptPointResult selection = acDoc.Editor.GetPoint("Выберите точку");
                    if (selection.Status == PromptStatus.OK)
                    {
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord acModSpace;
                        acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        lineTypeTable = (LinetypeTable)acTrans.GetObject(acDb.LinetypeTableId, OpenMode.ForRead);
                        if (!lineTypeTable.Has("штриховая2"))
                            try
                            {
                                acDb.LoadLineTypeFile("*", @"Data\acad.lin");
                                lineTypeTable = (LinetypeTable)acTrans.GetObject(acDb.LinetypeTableId, OpenMode.ForRead);
                            }
                            catch
                            {
                                editor.WriteMessage("В проекте не найден тип линий \"штриховая2\". Попытка загрузить файл с типом линии.. Не найден файл acad.lin.");
                            }

                        textStyle = (TextStyleTableRecord)acTrans.GetObject(acDb.Textstyle, OpenMode.ForWrite);
                        textStyle.FileName = @"Data\GOST_Common.ttf";
                        Point3d startPoint = selection.Value;
                        drawUnits(acTrans, acDb, acModSpace, units, drawSheet(acTrans, acModSpace, acDb, startPoint, firstSheet));
                        double lowestPoint = units[0].cableMarkLine.EndPoint.Y;
                        for (int i = 1; i < units.Count; i++)
                            if (units[i].cableMarkLine.EndPoint.Y < lowestPoint)
                                lowestPoint = units[i].cableMarkLine.EndPoint.Y;
                        for (int i = 0; i < units.Count; i++)
                        {
                            while (units[i].cableMarkLine.EndPoint.Y > lowestPoint)
                            {
                                units[i].cableMarkLine.EndPoint = units[i].cableMarkLine.EndPoint.Add(new Vector3d(0, -1, 0));
                                units[i].cableName.Location = units[i].cableMarkLine.EndPoint.Add(new Vector3d(-1, (units[i].cableMarkLine.StartPoint.Y - units[i].cableMarkLine.EndPoint.Y) / 2, 0));
                            }
                        }
                        int k = 0;
                        currentTable = tables[k];
                        for (int i = 0; i < units.Count; i++)
                        {
                            if (units[i].newSheet)
                                currentTable = tables[++k];
                            drawOther(acTrans, acModSpace, i);
                            currentUnit = units[i];
                        }
                        for (int i = 0; i < units.Count;)
                        {
                            List<int> group = findGroups(i);
                            if (group != null)
                            {
                                Line line = (Line)acTrans.GetObject(units[group[0]].outputCableLine.Id, OpenMode.ForRead);
                                double l = line.EndPoint.Y;
                                for (int j = 1; j < group.Count; j++)
                                {
                                    line = (Line)acTrans.GetObject(units[group[j]].outputCableLine.Id, OpenMode.ForRead);
                                    if (line.EndPoint.Y < lowestPoint)
                                    {
                                        l = line.EndPoint.Y;
                                    }
                                }
                                for (int j = 0; j < group.Count; j++)
                                {
                                    line = (Line)acTrans.GetObject(units[group[j]].outputCableLine.Id, OpenMode.ForWrite);
                                    while (line.EndPoint.Y > l)
                                        line.EndPoint = line.EndPoint.Add(new Vector3d(0, -1, 0));
                                }
                                i += group.Count;
                            }
                            else i++;
                        }
                        
                        acDoc.Editor.Regen();
                        currentTable = null;
                        currentTBox = null;
                        currentTBoxName = null;
                        currentLeader = null;
                        for (int i = 0; i < tBoxes.Count; i++)
                        {
                            if (tBoxes[i].Count == 1)
                            {
                                MText text = tBoxes[i].textName;
                                string str = text.Contents;
                                str = str.Remove(text.Contents.Length - 2);
                                text.Contents = str;

                                Leader ldr = tBoxes[i].ldr;
                                Point3d point = ldr.VertexAt(1);
                                ldr.SetVertexAt(2, new Point3d(point.X + text.ActualWidth + 1, point.Y, 0));
                            }
                        }
                        for (int i = 0; i < tables.Count; i++ )
                        {
                            if (!acBlkTbl.Has(tables[i].Id))
                            {
                                acModSpace.AppendEntity(tables[i]);
                                acTrans.AddNewlyCreatedDBObject(tables[i], true);
                            }
                        }
                        acDoc.Editor.Regen();
                        acTrans.Commit();
                        acTrans.Dispose();
                    }
                }
                acDb.Audit(true, true);
            }
        }

        private static Polyline drawSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d prevSheet, bool first)
        {
            Polyline shieldPoly = new Polyline();
            shieldPoly.SetDatabaseDefaults();
            shieldPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            shieldPoly.Closed = true;
            shieldPoly.AddVertexAt(0, new Point2d(prevSheet.X, prevSheet.Y), 0, 0, 0);
            shieldPoly.AddVertexAt(1, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(764, 0)), 0, 0, 0);
            shieldPoly.AddVertexAt(2, shieldPoly.GetPoint2dAt(1).Add(new Vector2d(0, -34)), 0, 0, 0);
            shieldPoly.AddVertexAt(3, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(0, -34)), 0, 0, 0);
            modSpace.AppendEntity(shieldPoly);
            acTrans.AddNewlyCreatedDBObject(shieldPoly, true);

            insertSheet(acTrans, modSpace, acdb, new Point3d(shieldPoly.GetPoint2dAt(0).X - 92, shieldPoly.GetPoint2dAt(0).Y + 10, 0), first);

            Polyline gndPoly = new Polyline();
            gndPoly.SetDatabaseDefaults();
            gndPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            gndPoly.Closed = true;
            gndPoly.AddVertexAt(0, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -21)), 0, 0, 0);
            gndPoly.AddVertexAt(1, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(754, -21)), 0, 0, 0);
            gndPoly.AddVertexAt(2, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(754, -23)), 0, 0, 0);
            gndPoly.AddVertexAt(3, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -23)), 0, 0, 0);
            modSpace.AppendEntity(gndPoly);
            acTrans.AddNewlyCreatedDBObject(gndPoly, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.TextStyleId = textStyle.Id;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = gndPoly.GetPoint3dAt(0).Add(new Vector3d(0, 2, 0));
            text.TextString = "Шина функционального заземления";
            text.HorizontalMode = TextHorizontalMode.TextLeft;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Table table = new Table();
            table.Position = shieldPoly.GetPoint3dAt(0).Add(new Vector3d(-30, -272, 0));
            table.SetSize(4, 1);
            table.TableStyle = acdb.Tablestyle;

            table.SetTextHeight(0, 0, 2.5);
            table.Cells[0, 0].TextString = "Тип оборудования";
            table.Cells[0, 0].TextStyleId = textStyle.Id;
            table.SetAlignment(0, 0, CellAlignment.MiddleCenter);
            table.Rows[0].IsMergeAllEnabled = false;

            table.SetTextHeight(1, 0, 2.5);
            table.Cells[1, 0].TextString = "Обозначение по проекту";
            table.Cells[1, 0].TextStyleId = textStyle.Id;
            table.SetAlignment(1, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(2, 0, 2.5);
            table.Cells[2, 0].TextString = "Параметры";
            table.Cells[2, 0].TextStyleId = textStyle.Id;
            table.SetAlignment(2, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(3, 0, 2.5);
            table.Cells[3, 0].TextString = "Оборудоание";
            table.Cells[3, 0].TextStyleId = textStyle.Id;
            table.SetAlignment(3, 0, CellAlignment.MiddleCenter);

            table.Columns[0].Width = 30;
            table.GenerateLayout();
            currentTable = table;
            tables.Add(table);

            return shieldPoly;
        }

        private static void drawUnits(Transaction acTrans, Database acdb, BlockTableRecord modSpace, List<Unit> units, Polyline shield)
        {
            Polyline tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[0].cupboardName);
            Polyline prevPoly = tBoxPoly;
            Polyline prevTermPoly = null;
            string prevTerminal=string.Empty;
            string prevCupboard = units[0].cupboardName;
            bool newJ = false;
            for (int j = 0; j < units.Count; j++)
            {
                cableLineUps = new List<Line>();
                cableLineText = new List<MText>();
                bool aborted = false;
                double leftEdgeX = 0;
                double rightEdgeX = 0;
                Point3d lowestPoint = Point3d.Origin;
                if (prevCupboard != units[j].cupboardName)
                {
                    firstSheet = false;
                    currentTBoxName = null;
                    shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)), firstSheet);
                    tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
                    prevPoly = tBoxPoly;
                    prevTermPoly = null;
                    prevTerminal = string.Empty;
                    prevCupboard = units[j].cupboardName;
                    newJ = false;
                    Unit unit = units[j];
                    unit.newSheet = true;
                    units[j] = unit;
                }
                using (Transaction trans = acdb.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < units[j].terminals.Count; i++)
                    {
                        string terminalTag = units[j].terminals[i].Split('[', ']')[1];
                        string terminal = units[j].terminals[i].Split('[', ']')[2];
                        if (prevTerminal == String.Empty)
                        {
                            prevTerminal = terminalTag;
                            DBText text = new DBText();
                            text.SetDatabaseDefaults();
                            text.TextStyleId = textStyle.Id;
                            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                            text.TextString = terminalTag;
                            text.HorizontalMode = TextHorizontalMode.TextCenter;
                            text.AlignmentPoint = text.Position;
                            modSpace.AppendEntity(text);
                            acTrans.AddNewlyCreatedDBObject(text, true);

                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal, 15, out lowestPoint, units[j].designation, i + 1);
                            leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                            newJ = false;
                        }
                        else if (prevTerminal == terminalTag)
                        {
                            Point2d p1;
                            Point2d p2;
                            if (newJ)
                            {
                                p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(60, 0));
                                prevPoly.SetPointAt(1, p1);
                                p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(60, 0));
                                prevPoly.SetPointAt(2, p2);
                            }
                            else
                            {
                                p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(6, 0));
                                prevPoly.SetPointAt(1, p1);
                                p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(6, 0));
                                prevPoly.SetPointAt(2, p2);
                            }
                            if (p1.X >= shield.GetPoint2dAt(1).X - 170)
                            {
                                trans.Abort();
                                currentTBoxName = null;
                                aborted = true;
                                firstSheet = false;
                                shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950,0,0)), firstSheet);
                                tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[0].cupboardName);
                                prevPoly = tBoxPoly;
                                prevTermPoly = null;
                                prevTerminal = string.Empty;
                                newJ = false;
                                Unit unit = units[j];
                                unit.newSheet = true;
                                units[j] = unit;
                                break;
                            }
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevTermPoly, terminal, newJ?60:6, out lowestPoint, units[j].designation, i+1);
                            if (newJ) newJ = false;
                            if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                        }
                        else
                        {
                            newJ = false;
                            tBoxPoly = new Polyline();
                            tBoxPoly.SetDatabaseDefaults();
                            tBoxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            tBoxPoly.Closed = true;
                            tBoxPoly.AddVertexAt(0, prevPoly.GetPoint2dAt(1).Add(new Vector2d(40, 0)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(1, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(21, 0)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(2, tBoxPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(3, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
                            modSpace.AppendEntity(tBoxPoly);
                            acTrans.AddNewlyCreatedDBObject(tBoxPoly, true);
                            prevPoly = tBoxPoly;
                            prevTermPoly = null;
                            prevTerminal = terminalTag;

                            DBText text = new DBText();
                            text.SetDatabaseDefaults();
                            text.TextStyleId = textStyle.Id;
                            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                            text.TextString = terminalTag;
                            text.HorizontalMode = TextHorizontalMode.TextCenter;
                            text.AlignmentPoint = text.Position;
                            modSpace.AppendEntity(text);
                            acTrans.AddNewlyCreatedDBObject(text, true);

                            if (tBoxPoly.GetPoint2dAt(1).X >= shield.GetPoint2dAt(1).X - 170) 
                            {
                                trans.Abort();
                                currentTBoxName = null;
                                aborted = true;
                                firstSheet = false;
                                shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)), firstSheet);
                                tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[0].cupboardName);
                                prevPoly = tBoxPoly;
                                prevTermPoly = null;
                                prevTerminal = string.Empty;
                                newJ = false;
                                Unit unit = units[j];
                                unit.newSheet = true;
                                units[j] = unit;
                                break;
                            }
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal, 15, out lowestPoint, units[j].designation, i+1);
                            double x = prevPoly.GetPoint2dAt(1).X;
                            double xx = prevTermPoly.GetPoint2dAt(1).X;
                            if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                        }
                    }
                    if (!aborted)
                    {
                        rightEdgeX = prevTermPoly.GetPoint2dAt(1).X;
                        if (units[j].shield)
                        {
                            Polyline groundLasso = new Polyline();
                            groundLasso.SetDatabaseDefaults();
                            if (lineTypeTable.Has("штриховая2"))
                                groundLasso.LinetypeId = lineTypeTable["штриховая2"];
                            else if (lineTypeTable.Has("hidden2"))
                                groundLasso.LinetypeId = lineTypeTable["hidden2"];
                            groundLasso.LinetypeScale = 5;
                            groundLasso.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            groundLasso.Closed = true;
                            groundLasso.AddVertexAt(0, new Point2d(leftEdgeX + 3, prevTermPoly.GetPoint2dAt(2).Y - 3), 0, 0, 0);
                            groundLasso.AddVertexAt(1, new Point2d(rightEdgeX - 3, prevTermPoly.GetPoint2dAt(2).Y - 3), -1, 0, 0);
                            groundLasso.AddVertexAt(2, groundLasso.GetPoint2dAt(1).Add(new Vector2d(0, -3.28)), 0, 0, 0);
                            groundLasso.AddVertexAt(3, groundLasso.GetPoint2dAt(0).Add(new Vector2d(0, -3.28)), -1, 0, 0);
                            modSpace.AppendEntity(groundLasso);
                            acTrans.AddNewlyCreatedDBObject(groundLasso, true);

                            Line groundLine1 = new Line();
                            groundLine1.SetDatabaseDefaults();
                            if (lineTypeTable.Has("штриховая2"))
                                groundLine1.LinetypeId = lineTypeTable["штриховая2"];
                            else if (lineTypeTable.Has("hidden2"))
                                groundLine1.LinetypeId = lineTypeTable["hidden2"];
                            groundLine1.LinetypeScale = 5;
                            groundLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            groundLine1.StartPoint = groundLasso.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
                            groundLine1.EndPoint = groundLine1.StartPoint.Add(new Vector3d(6, 0, 0));
                            modSpace.AppendEntity(groundLine1);
                            acTrans.AddNewlyCreatedDBObject(groundLine1, true);

                            Circle acCirc = new Circle();
                            acCirc.SetDatabaseDefaults();
                            if (lineTypeTable.Has("штриховая2"))
                                acCirc.LinetypeId = lineTypeTable["штриховая2"];
                            else if (lineTypeTable.Has("hidden2"))
                                acCirc.LinetypeId = lineTypeTable["hidden2"];
                            acCirc.LinetypeScale = 5;
                            acCirc.Center = groundLasso.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
                            acCirc.Radius = 0.37;
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

                            Line groundLine2 = new Line();
                            groundLine2.SetDatabaseDefaults();
                            if (lineTypeTable.Has("штриховая2"))
                                groundLine2.LinetypeId = lineTypeTable["штриховая2"];
                            else if (lineTypeTable.Has("hidden2"))
                                groundLine2.LinetypeId = lineTypeTable["hidden2"];
                            groundLine2.LinetypeScale = 5;
                            groundLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            groundLine2.StartPoint = groundLine1.EndPoint;
                            groundLine2.EndPoint = groundLine2.StartPoint.Add(new Vector3d(0, shield.StartPoint.Y - groundLine2.StartPoint.Y - 22.5, 0));
                            modSpace.AppendEntity(groundLine2);
                            acTrans.AddNewlyCreatedDBObject(groundLine2, true);

                            Circle groundCircle = new Circle();
                            groundCircle.SetDatabaseDefaults();
                            if (lineTypeTable.Has("штриховая2"))
                                groundCircle.LinetypeId = lineTypeTable["штриховая2"];
                            else if (lineTypeTable.Has("hidden2"))
                                groundCircle.LinetypeId = lineTypeTable["hidden2"];
                            groundCircle.LinetypeScale = 5;
                            groundCircle.Center = new Point3d(groundLine2.EndPoint.X, groundLine2.EndPoint.Y + 0.36, 0);
                            groundCircle.Radius = 0.36;
                            modSpace.AppendEntity(groundCircle);
                            acTrans.AddNewlyCreatedDBObject(groundCircle, true);
                        }
                        units[j] = drawCable(acTrans, modSpace, new Point3d(leftEdgeX + 3, lowestPoint.Y, 0), new Point3d(rightEdgeX - 3, lowestPoint.Y, 0), units[j]);
                        trans.Commit();
                    }
                    else j--;
                }
                newJ = true;
            }
        }

        private static Polyline drawShieldTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Polyline shield, string cupbordName)
        {
            Polyline boxPoly = new Polyline();
            boxPoly.SetDatabaseDefaults();
            boxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            boxPoly.Closed = true;
            boxPoly.AddVertexAt(0, shield.GetPoint2dAt(0).Add(new Vector2d(5, -25)), 0, 0, 0);
            boxPoly.AddVertexAt(1, boxPoly.GetPoint2dAt(0).Add(new Vector2d(21, 0)), 0, 0, 0);
            boxPoly.AddVertexAt(2, boxPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            boxPoly.AddVertexAt(3, boxPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(boxPoly);
            acTrans.AddNewlyCreatedDBObject(boxPoly, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.TextStyleId = textStyle.Id;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = shield.GetPoint3dAt(0).Add(new Vector3d(382, -16, 0));
            text.TextString = cupbordName;
            text.Height = 4;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            return boxPoly;
        }

        private static Polyline drawTerminal(Transaction acTrans, BlockTableRecord modSpace, Polyline prevPoly, string terminal, double offset, out Point3d lowestPoint, string cableMark, int cableNumber)
        {
            Polyline termPoly = new Polyline();
            termPoly.SetDatabaseDefaults();
            termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly.Closed = true;
            termPoly.AddVertexAt(0, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), 0, 0, 0);
            termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(6, 0)), 0, 0, 0);
            termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(termPoly);
            acTrans.AddNewlyCreatedDBObject(termPoly, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.TextStyleId = textStyle.Id;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = termPoly.GetPoint3dAt(0).Add(new Vector3d(3, -4, 0));
            text.TextString = terminal;
            text.Rotation = 1.5708;
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Line cableLineUp = new Line();
            cableLineUp.SetDatabaseDefaults();
            cableLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp.StartPoint = new Point3d(termPoly.GetPoint2dAt(3).X+(termPoly.GetPoint2dAt(2).X - termPoly.GetPoint2dAt(3).X) / 2, termPoly.GetPoint2dAt(3).Y, 0);
            cableLineUp.EndPoint = cableLineUp.StartPoint.Add(new Vector3d(0, -44, 0));
            modSpace.AppendEntity(cableLineUp);
            acTrans.AddNewlyCreatedDBObject(cableLineUp, true);
           
            MText textLine = new MText();
            textLine.SetDatabaseDefaults();
            textLine.TextStyleId = textStyle.Id;
            textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            textLine.Location = cableLineUp.EndPoint.Add(new Vector3d(-1, 2, 0));
            textLine.Contents = cableMark + "-" + cableNumber;
            textLine.Rotation = 1.5708;
            textLine.Attachment = AttachmentPoint.BottomLeft;
            modSpace.AppendEntity(textLine);
            acTrans.AddNewlyCreatedDBObject(textLine, true);

            while (cableLineUp.Length - 8 < textLine.ActualWidth)
            {
                cableLineUp.EndPoint = cableLineUp.EndPoint.Add(new Vector3d(0, -1, 0));
                textLine.Location = cableLineUp.EndPoint.Add(new Vector3d(-1, 2, 0));
            }
            cableLineUps.Add(cableLineUp);
            cableLineText.Add(textLine);
            double lowestpoint = cableLineUps.Count>0?cableLineUps[0].EndPoint.Y:cableLineUp.EndPoint.Y;
            for (int i = 1; i < cableLineUps.Count; i++)
            {
                if (cableLineUps[i].EndPoint.Y<lowestpoint)
                    lowestpoint = cableLineUps[i].EndPoint.Y;
            }
            for (int i = 0; i < cableLineUps.Count; i++)
            {
                while (cableLineUps[i].EndPoint.Y - 8 > lowestpoint)
                {
                    cableLineUps[i].EndPoint = cableLineUps[i].EndPoint.Add(new Vector3d(0, -1, 0));
                    cableLineText[i].Location = cableLineUps[i].EndPoint.Add(new Vector3d(-1, 2, 0));
                }
            }
               
            lowestPoint = cableLineUp.EndPoint;
            return termPoly;
        }

        private static Unit drawCable(Transaction acTrans, BlockTableRecord modSpace, Point3d firstCable, Point3d lastCable, Unit unit)
        {
            Line jumperLineUp = new Line();
            jumperLineUp.SetDatabaseDefaults();
            jumperLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            jumperLineUp.StartPoint = firstCable;
            jumperLineUp.EndPoint = lastCable;
            modSpace.AppendEntity(jumperLineUp);
            acTrans.AddNewlyCreatedDBObject(jumperLineUp, true);

            Line cableLine1 = new Line();
            cableLine1.SetDatabaseDefaults();
            cableLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            Point3d startPoint = new Point3d(jumperLineUp.StartPoint.X + (jumperLineUp.EndPoint.X - jumperLineUp.StartPoint.X) / 2, jumperLineUp.StartPoint.Y, 0);
            Point3d endPoint = new Point3d(startPoint.X, startPoint.Y - 5, 0);
            cableLine1.StartPoint = startPoint;
            cableLine1.EndPoint = endPoint;
            modSpace.AppendEntity(cableLine1);
            acTrans.AddNewlyCreatedDBObject(cableLine1, true);

            Polyline cableName = new Polyline();
            cableName.SetDatabaseDefaults();
            cableName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableName.Closed = true;
            cableName.AddVertexAt(0, new Point2d(cableLine1.EndPoint.X - 12.5, cableLine1.EndPoint.Y), 0, 0, 0);
            cableName.AddVertexAt(1, cableName.GetPoint2dAt(0).Add(new Vector2d(25, 0)), -1, 0, 0);
            cableName.AddVertexAt(2, cableName.GetPoint2dAt(1).Add(new Vector2d(0, -10)), 0, 0, 0);
            cableName.AddVertexAt(3, cableName.GetPoint2dAt(0).Add(new Vector2d(0, -10)), -1, 0, 0);
            modSpace.AppendEntity(cableName);
            acTrans.AddNewlyCreatedDBObject(cableName, true);

            double X = cableName.GetPoint2dAt(0).X + (cableName.GetPoint2dAt(1).X - cableName.GetPoint2dAt(0).X)/2;
            double Y = cableName.GetPoint2dAt(0).Y - 2;
            Point3d cableNameCentre = new Point3d(X, Y, 0);
            MText textName = new MText();
            textName.SetDatabaseDefaults();
            textName.TextStyleId = textStyle.Id;
            textName.TextStyleId = textStyle.Id;
            textName.Location = cableNameCentre;
            textName.Attachment = AttachmentPoint.MiddleCenter;
            textName.Width = 25;
            textName.Contents = unit.tBoxName == string.Empty ? unit.cupboardName.Split(' ')[1] + "/" + unit.designation : unit.cupboardName.Split(' ')[1] + "/" + unit.tBoxName;
            modSpace.AppendEntity(textName);
            acTrans.AddNewlyCreatedDBObject(textName, true);
            if (textName.Width < textName.ActualWidth)
            {
                textName.Contents = unit.tBoxName == string.Empty ? unit.cupboardName.Split(' ')[1] + "/ " + unit.designation : unit.cupboardName.Split(' ')[1] + "/ " + unit.tBoxName;
                textName.Attachment = AttachmentPoint.TopCenter;
                while (textName.ActualHeight > cableName.GetPoint2dAt(0).Y - cableName.GetPoint2dAt(3).Y - 4)
                {
                    cableName.SetPointAt(2, cableName.GetPoint2dAt(2).Add(new Vector2d(0, -1)));
                    cableName.SetPointAt(3, cableName.GetPoint2dAt(3).Add(new Vector2d(0, -1)));
                }
            }

            Line cableLine2 = new Line();
            cableLine2.SetDatabaseDefaults();
            cableLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            startPoint = new Point3d(endPoint.X, cableName.GetPoint2dAt(2).Y, 0);
            double cableLine2Length = unit.tBoxName != string.Empty ? startPoint.Y - 40 : startPoint.Y - 80;
            endPoint = new Point3d(startPoint.X, cableLine2Length, 0);
            cableLine2.StartPoint = startPoint;
            cableLine2.EndPoint = endPoint;
            modSpace.AppendEntity(cableLine2);
            acTrans.AddNewlyCreatedDBObject(cableLine2, true);

            MText cableMark = new MText();
            cableMark.SetDatabaseDefaults();
            cableMark.TextStyleId = textStyle.Id;
            cableMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableMark.Location = cableLine2.EndPoint.Add(new Vector3d(-1, (cableLine2.StartPoint.Y - cableLine2.EndPoint.Y)/2, 0));
            cableMark.Contents = unit.cableMark;
            cableMark.Rotation = 1.5708;
            cableMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(cableMark);
            acTrans.AddNewlyCreatedDBObject(cableMark, true);

            while (cableLine2.Length - 8 < cableMark.ActualWidth)
            {
                cableLine2.EndPoint = cableLine2.EndPoint.Add(new Vector3d(0, -1, 0));
                cableMark.Location = cableLine2.EndPoint.Add(new Vector3d(-1, (cableLine2.StartPoint.Y - cableLine2.EndPoint.Y)/2, 0));
            }

            unit.cableMarkLine = cableLine2;
            unit.cableName = cableMark;
            return unit;
        }

        private static Polyline drawOther(Transaction acTrans, BlockTableRecord modSpace, int ind)
        {
            double leftEdgeX = 0;
            double rightEdgeX = 0;
            Line cableLine2 = units[ind].cableMarkLine;

            Point3d point;
            if (units[ind].tBoxName != string.Empty)
            {
                point = drawTerminalBox(acTrans, modSpace, cableLine2, ind);
            }
            else
            {
                point = cableLine2.EndPoint;
                currentTBoxName = null;
            }
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = point.X - ((units[ind].equipTerminals.Count - 1) * 5) / 2;
            jumperLineDown.StartPoint = new Point3d(x, point.Y, 0);
            x = x + ((units[ind].equipTerminals.Count - 1) * 5);
            jumperLineDown.EndPoint = new Point3d(x, point.Y, 0);
            modSpace.AppendEntity(jumperLineDown);
            acTrans.AddNewlyCreatedDBObject(jumperLineDown, true);

            List<Line> lines = new List<Line>();
            List<MText> texts = new List<MText>();
            double lowestPoint = 0;
            for (int i = 0; i < units[ind].equipTerminals.Count; i++)
            {
                Line cableLineDown = new Line();
                cableLineDown.SetDatabaseDefaults();
                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableLineDown.StartPoint = jumperLineDown.StartPoint.Add(new Vector3d(i * 5, 0, 0));
                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -44, 0));
                modSpace.AppendEntity(cableLineDown);
                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                MText textLine = new MText();
                textLine.SetDatabaseDefaults();
                textLine.TextStyleId = textStyle.Id;
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Location = cableLineDown.EndPoint.Add(new Vector3d(-1, 1, 0));
                textLine.Contents = units[ind].designation + "-" + (i + 1);
                textLine.Rotation = 1.5708;
                textLine.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);

                lines.Add(cableLineDown);
                texts.Add(textLine);

                while (textLine.ActualWidth > cableLineDown.StartPoint.Y - cableLineDown.EndPoint.Y - 7)
                {
                    cableLineDown.EndPoint = cableLineDown.EndPoint.Add(new Vector3d(0, -1, 0));
                    textLine.Location = cableLineDown.EndPoint.Add(new Vector3d(-1, 1, 0));
                }
                if (i == 0) lowestPoint = cableLineDown.EndPoint.Y;
                else if (lowestPoint > cableLineDown.EndPoint.Y)
                    lowestPoint = cableLineDown.EndPoint.Y;

                if (i == 0) leftEdgeX = cableLineDown.StartPoint.X - 3;
                rightEdgeX = cableLineDown.StartPoint.X + 3;
            }

            for (int i = 0; i < units[ind].equipTerminals.Count; i++)
            {
                Line cableLineDown = lines[i];
                MText textLine = texts[i];

                while (lowestPoint < cableLineDown.EndPoint.Y)
                {
                    cableLineDown.EndPoint = cableLineDown.EndPoint.Add(new Vector3d(0, -1, 0));
                    textLine.Location = cableLineDown.EndPoint.Add(new Vector3d(-1, 1, 0));
                }

                Polyline terminal = new Polyline();
                terminal.SetDatabaseDefaults();
                terminal.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                terminal.Closed = true;
                terminal.AddVertexAt(0, new Point2d(cableLineDown.EndPoint.X - 2.75, cableLineDown.EndPoint.Y), 0, 0, 0);
                terminal.AddVertexAt(1, terminal.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                terminal.AddVertexAt(2, terminal.GetPoint2dAt(1).Add(new Vector2d(0, -5)), 0, 0, 0);
                terminal.AddVertexAt(3, terminal.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                modSpace.AppendEntity(terminal);
                acTrans.AddNewlyCreatedDBObject(terminal, true);

                DBText text = new DBText();
                text.SetDatabaseDefaults();
                text.TextStyleId = textStyle.Id;
                text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                text.Position = terminal.GetPoint3dAt(0).Add(new Vector3d(2.5, -3.5, 0));
                text.TextString = units[ind].equipTerminals[i];
                text.HorizontalMode = TextHorizontalMode.TextCenter;
                text.AlignmentPoint = text.Position;
                modSpace.AppendEntity(text);
                acTrans.AddNewlyCreatedDBObject(text, true);
            }

            if (units[ind].shield)
            {
                Polyline groundLasso = new Polyline();
                groundLasso.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    groundLasso.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    groundLasso.LinetypeId = lineTypeTable["hidden2"];
                groundLasso.LinetypeScale = 5;
                groundLasso.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLasso.Closed = true;
                groundLasso.AddVertexAt(0, new Point2d(leftEdgeX + 3, jumperLineDown.StartPoint.Y - 3), 0, 0, 0);
                groundLasso.AddVertexAt(1, new Point2d(rightEdgeX - 3, jumperLineDown.StartPoint.Y - 3), -1, 0, 0);
                groundLasso.AddVertexAt(2, groundLasso.GetPoint2dAt(1).Add(new Vector2d(0, -3.28)), 0, 0, 0);
                groundLasso.AddVertexAt(3, groundLasso.GetPoint2dAt(0).Add(new Vector2d(0, -3.28)), -1, 0, 0);
                modSpace.AppendEntity(groundLasso);
                acTrans.AddNewlyCreatedDBObject(groundLasso, true);
            }
            Polyline equip = new Polyline();
            equip.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                equip.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                equip.LinetypeId = lineTypeTable["hidden2"];
            equip.LinetypeScale = 5;
            equip.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            equip.Closed = true;
            equip.AddVertexAt(0, new Point2d(jumperLineDown.StartPoint.X - 10, jumperLineDown.StartPoint.Y + 3.5), 0, 0, 0);
            equip.AddVertexAt(1, new Point2d(jumperLineDown.EndPoint.X + 10, jumperLineDown.StartPoint.Y + 3.5), 0, 0, 0);
            equip.AddVertexAt(2, new Point2d(jumperLineDown.EndPoint.X + 10, lowestPoint - 8.5), 0, 0, 0);
            equip.AddVertexAt(3, new Point2d(jumperLineDown.StartPoint.X - 10, lowestPoint - 8.5), 0, 0, 0);
            modSpace.AppendEntity(equip);
            acTrans.AddNewlyCreatedDBObject(equip, true);

            while (equip.GetPoint2dAt(3).Y <= currentTable.Position.Y)
                currentTable.Position = currentTable.Position.Add(new Vector3d(0, -1, 0));
            currentTable.Position = currentTable.Position.Add(new Vector3d(0, -2, 0));

            double C = equip.GetPoint2dAt(0).X;
            double D = equip.GetPoint2dAt(1).X;
            double A = currentTable.Position.X;
            double B = currentTable.Position.X + currentTable.Width;
            width = 2 * (C - B) + (D - C);
            currentTable.InsertColumns(currentTable.Columns.Count, width, 1);
            currentTable.UnmergeCells(currentTable.Rows[0]);
            currentTable.Columns[currentTable.Columns.Count - 1].TextStyleId = textStyle.Id;
            currentTable.Columns[currentTable.Columns.Count - 1].TextHeight = 2.5;
            currentTable.Cells[0, currentTable.Columns.Count - 1].TextString = units[ind].equipType == "" ? "-" : units[ind].equipType;
            currentTable.Cells[1, currentTable.Columns.Count - 1].TextString = units[ind].designation;
            currentTable.Cells[2, currentTable.Columns.Count - 1].TextString = units[ind].param;
            currentTable.Cells[3, currentTable.Columns.Count - 1].TextString = units[ind].equipment;
            currentTable.GenerateLayout();

            return equip;
        }

        private static Point3d drawTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Line cableLine, int ind)
        {
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            int count = units[ind].shield ? units[ind].equipTerminals.Count : units[ind].equipTerminals.Count - 1;
            double x = cableLine.EndPoint.X - count * 5 / 2;
            jumperLineDown.StartPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            x = x + count * 5;
            jumperLineDown.EndPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            modSpace.AppendEntity(jumperLineDown);
            acTrans.AddNewlyCreatedDBObject(jumperLineDown, true);

            double lowestPointY = 0;
            double leftEdgeX = 0;
            double rightEdgeX = 0;
            tBoxUnit tBox;
            int index = 0;
            bool exist = false;
            if (tBoxes.Exists(tbox => tbox.Name == units[ind].tBoxName))
            {
                index = tBoxes.IndexOf(tBoxes.Find(tbox => tbox.Name == units[ind].tBoxName));
                tBox = tBoxes[index];
                exist = true;
            }
            else tBox = new tBoxUnit(units[ind].tBoxName, 0, 1);
            double lowestPoint = 0;
            List<Polyline> polys = new List<Polyline>();
            List<MText> texts = new List<MText>();
            for (int i = 0; i < units[ind].equipTerminals.Count; i++)
            {
                Line cableLineInput = new Line();
                cableLineInput.SetDatabaseDefaults();
                cableLineInput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableLineInput.StartPoint = new Point3d(jumperLineDown.StartPoint.X + i * 5, jumperLineDown.StartPoint.Y, 0);
                cableLineInput.EndPoint = new Point3d(cableLineInput.StartPoint.X, cableLineInput.StartPoint.Y - 8, 0);
                modSpace.AppendEntity(cableLineInput);
                acTrans.AddNewlyCreatedDBObject(cableLineInput, true);

                Polyline terminal = new Polyline();
                terminal.SetDatabaseDefaults();
                terminal.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                terminal.Closed = true;
                terminal.AddVertexAt(0, new Point2d(cableLineInput.EndPoint.X - 2.75, cableLineInput.EndPoint.Y), 0, 0, 0);
                terminal.AddVertexAt(1, terminal.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                terminal.AddVertexAt(2, terminal.GetPoint2dAt(1).Add(new Vector2d(0, -34)), 0, 0, 0);
                terminal.AddVertexAt(3, terminal.GetPoint2dAt(0).Add(new Vector2d(0, -34)), 0, 0, 0);
                modSpace.AppendEntity(terminal);
                acTrans.AddNewlyCreatedDBObject(terminal, true);
                
                MText textLine = new MText();
                textLine.SetDatabaseDefaults();
                textLine.TextStyleId = textStyle.Id;
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                textLine.Contents = units[ind].designation + "-" + (i + 1);
                textLine.Rotation = 1.5708;
                textLine.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);
                
                polys.Add(terminal);
                texts.Add(textLine);

                while (textLine.ActualWidth > terminal.GetPoint2dAt(1).Y - terminal.GetPoint2dAt(2).Y - 7)
                {
                    terminal.SetPointAt(2, terminal.GetPoint2dAt(2).Add(new Vector2d(0, -1)));
                    terminal.SetPointAt(3, terminal.GetPoint2dAt(3).Add(new Vector2d(0, -1)));
                    textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                }
                if (i == 0) lowestPoint = terminal.GetPoint2dAt(2).Y;
                else if (lowestPoint > terminal.GetPoint2dAt(2).Y) 
                    lowestPoint = terminal.GetPoint2dAt(2).Y;
            }

            for (int i = 0; i < units[ind].equipTerminals.Count; i++)
            {
                Polyline terminal = polys[i];
                MText textLine = texts[i];

                while (lowestPoint < terminal.GetPoint2dAt(2).Y)
                {
                    terminal.SetPointAt(2, terminal.GetPoint2dAt(2).Add(new Vector2d(0, -1)));
                    terminal.SetPointAt(3, terminal.GetPoint2dAt(3).Add(new Vector2d(0, -1)));
                    textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                }

                Line line = new Line();
                line.SetDatabaseDefaults();
                line.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line.StartPoint = terminal.GetPoint3dAt(3).Add(new Vector3d(0, 5, 0));
                line.EndPoint = terminal.GetPoint3dAt(2).Add(new Vector3d(0, 5, 0));
                modSpace.AppendEntity(line);
                acTrans.AddNewlyCreatedDBObject(line, true);

                DBText text = new DBText();
                text.SetDatabaseDefaults();
                text.TextStyleId = textStyle.Id;
                text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                text.Position = line.StartPoint.Add(new Vector3d(2.5, -3.5, 0));
                text.TextString = tBox.LastTerminalNumber.ToString();
                text.HorizontalMode = TextHorizontalMode.TextCenter;
                text.AlignmentPoint = text.Position;
                modSpace.AppendEntity(text);
                acTrans.AddNewlyCreatedDBObject(text, true);
                tBox.LastTerminalNumber++;

                Line cableLineOutput = new Line();
                cableLineOutput.SetDatabaseDefaults();
                cableLineOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableLineOutput.StartPoint = new Point3d(jumperLineDown.StartPoint.X + i * 5, terminal.GetPoint2dAt(2).Y, 0);
                cableLineOutput.EndPoint = cableLineOutput.StartPoint.Add(new Vector3d(0, -8, 0));
                modSpace.AppendEntity(cableLineOutput);
                acTrans.AddNewlyCreatedDBObject(cableLineOutput, true);

                lowestPointY = cableLineOutput.EndPoint.Y;
                if (i == 0) leftEdgeX = cableLineOutput.StartPoint.X - 3;
                rightEdgeX = cableLineOutput.StartPoint.X + 3;
            }

            if (units[ind].shield)
            {
                Polyline terminal = new Polyline();
                terminal.SetDatabaseDefaults();
                terminal.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                terminal.Closed = true;
                terminal.AddVertexAt(0, new Point2d(rightEdgeX - 0.75, jumperLineDown.EndPoint.Y - 8), 0, 0, 0);
                terminal.AddVertexAt(1, terminal.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                terminal.AddVertexAt(2, new Point2d(terminal.GetPoint2dAt(0).X+5, lowestPoint), 0, 0, 0);
                terminal.AddVertexAt(3, new Point2d(terminal.GetPoint2dAt(0).X, lowestPoint), 0, 0, 0);
                modSpace.AppendEntity(terminal);
                acTrans.AddNewlyCreatedDBObject(terminal, true);

                MText textLine = new MText();
                textLine.SetDatabaseDefaults();
                textLine.TextStyleId = textStyle.Id;
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                textLine.Contents = "шина ф. заземл.";
                textLine.Rotation = 1.5708;
                textLine.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);

                Line line = new Line();
                line.SetDatabaseDefaults();
                line.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line.StartPoint = terminal.GetPoint3dAt(3).Add(new Vector3d(0, 5, 0));
                line.EndPoint = terminal.GetPoint3dAt(2).Add(new Vector3d(0, 5, 0));
                modSpace.AppendEntity(line);
                acTrans.AddNewlyCreatedDBObject(line, true);

                jumperLineDown.EndPoint = jumperLineDown.EndPoint.Add(new Vector3d(-5, 0, 0));

                drawGndLasso(acTrans, modSpace, leftEdgeX, rightEdgeX, terminal.GetPoint2dAt(0).X, terminal.GetPoint2dAt(1).X, terminal.GetPoint2dAt(0).Y, true);
                drawGndLasso(acTrans, modSpace, leftEdgeX, rightEdgeX, terminal.GetPoint2dAt(3).X, terminal.GetPoint2dAt(2).X, terminal.GetPoint2dAt(3).Y, false);
            }
            
            Line jumperOutput = new Line();
            jumperOutput.SetDatabaseDefaults();
            jumperOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            jumperOutput.StartPoint = new Point3d(jumperLineDown.StartPoint.X, lowestPointY, 0);
            jumperOutput.EndPoint = new Point3d(jumperLineDown.EndPoint.X, lowestPointY, 0);
            modSpace.AppendEntity(jumperOutput);
            acTrans.AddNewlyCreatedDBObject(jumperOutput, true);

            Line cLineOutput = new Line();
            cLineOutput.SetDatabaseDefaults();
            cLineOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cLineOutput.StartPoint = new Point3d(cableLine.EndPoint.X, jumperOutput.EndPoint.Y, 0);
            cLineOutput.EndPoint = cLineOutput.StartPoint.Add(new Vector3d(0, -12, 0));
            modSpace.AppendEntity(cLineOutput);
            acTrans.AddNewlyCreatedDBObject(cLineOutput, true);
            Unit unit = units[ind];
            unit.outputCableLine = cLineOutput;
            units[ind] = unit;

            if (currentTBoxName == null || currentUnit.tBoxName != units[ind].tBoxName)
            {
                Polyline tbox = new Polyline();
                tbox.SetDatabaseDefaults();
                tbox.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tbox.Closed = true;
                tbox.AddVertexAt(0, new Point2d(jumperLineDown.StartPoint.X - 7, jumperLineDown.StartPoint.Y + 2), 0, 0, 0);
                tbox.AddVertexAt(1, new Point2d(jumperLineDown.EndPoint.X + 9, jumperLineDown.StartPoint.Y + 2), 0, 0, 0);
                tbox.AddVertexAt(2, new Point2d(jumperLineDown.EndPoint.X + 9, jumperOutput.StartPoint.Y - 2), 0, 0, 0);
                tbox.AddVertexAt(3, new Point2d(jumperLineDown.StartPoint.X - 7, jumperOutput.StartPoint.Y - 2), 0, 0, 0);
                modSpace.AppendEntity(tbox);
                acTrans.AddNewlyCreatedDBObject(tbox, true);
                currentTBox = tbox;

                DBText tBoxText = new DBText();
                tBoxText.SetDatabaseDefaults();
                tBoxText.TextStyleId = textStyle.Id;
                tBoxText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxText.Position = tbox.GetPoint3dAt(0).Add(new Vector3d(4, -28, 0));
                tBoxText.TextString = "XT";
                tBoxText.Rotation = 1.5708;
                tBoxText.VerticalMode = TextVerticalMode.TextBottom;
                tBoxText.HorizontalMode = TextHorizontalMode.TextCenter;
                tBoxText.AlignmentPoint = tBoxText.Position;
                modSpace.AppendEntity(tBoxText);
                acTrans.AddNewlyCreatedDBObject(tBoxText, true);

                tBox.Count++;
                MText tBoxName = new MText();
                tBoxName.SetDatabaseDefaults();
                tBoxName.TextStyleId = textStyle.Id;
                tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxName.Location = tbox.GetPoint3dAt(1).Add(new Vector3d(1, 6, 0));
                tBoxName.Contents = tBox.Name+"."+tBox.Count.ToString();
                tBoxName.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(tBoxName);
                acTrans.AddNewlyCreatedDBObject(tBoxName, true);
                currentTBoxName = tBoxName;
                tBox.textName = tBoxName;

                Leader acLdr = new Leader();
                acLdr.SetDatabaseDefaults();
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(tBoxName.ActualWidth+1, 5, 0)));
                acLdr.HasArrowHead = false;
                tBox.ldr = acLdr;

                modSpace.AppendEntity(acLdr);
                acTrans.AddNewlyCreatedDBObject(acLdr, true);
                currentLeader = acLdr;
            }
            else
            {
                currentTBox.SetPointAt(1, new Point2d(jumperLineDown.EndPoint.X + 9, jumperLineDown.StartPoint.Y + 2));
                currentTBox.SetPointAt(2, new Point2d(jumperLineDown.EndPoint.X + 9, jumperOutput.StartPoint.Y - 2));
                if (currentTBox.GetPoint2dAt(3).Y > currentTBox.GetPoint2dAt(2).Y)
                    currentTBox.SetPointAt(3, new Point2d(currentTBox.GetPoint2dAt(3).X, currentTBox.GetPoint2dAt(2).Y));
                else
                    currentTBox.SetPointAt(2, new Point2d(currentTBox.GetPoint2dAt(2).X, currentTBox.GetPoint2dAt(3).Y));

                currentTBoxName.Erase();
                MText tBoxName = new MText();
                tBoxName.SetDatabaseDefaults();
                tBoxName.TextStyleId = textStyle.Id;
                tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxName.Location = currentTBox.GetPoint3dAt(1).Add(new Vector3d(1, 6, 0));
                tBoxName.Contents = tBox.Name + "." + tBox.Count.ToString();
                tBoxName.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(tBoxName);
                acTrans.AddNewlyCreatedDBObject(tBoxName, true);
                currentTBoxName = tBoxName;
                tBox.textName = tBoxName;

                currentLeader.Erase();
                Leader acLdr = new Leader();
                acLdr.SetDatabaseDefaults();
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(tBoxName.ActualWidth+1, 5, 0)));
                acLdr.HasArrowHead = false;
                modSpace.AppendEntity(acLdr);
                acTrans.AddNewlyCreatedDBObject(acLdr, true);
                currentLeader = acLdr;
                tBox.ldr = acLdr;
            }

            if (exist) tBoxes[index] = tBox;
            else tBoxes.Add(tBox);

            return cLineOutput.EndPoint;
        }

        private static void insertSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d point, bool first)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename;
            string blockName;
            if (first)
            {
                filename = @"Data\Frame.dwg";
                blockName = "Frame";
            }
            else
            {
                filename = @"Data\FrameOther.dwg";
                blockName = "FrameOther";
            }
            using(Database sourceDb = new Database(false, true))
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
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void drawGndLasso(Transaction acTrans, BlockTableRecord modSpace, double leftEdgeX, double rightEdgeX, double lPoint, double rPoint, double yPoint, bool Upper)
        {
            Polyline groundLasso = new Polyline();
            groundLasso.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLasso.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLasso.LinetypeId = lineTypeTable["hidden2"];
            groundLasso.LinetypeScale = 5;
            groundLasso.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLasso.Closed = true;
            groundLasso.AddVertexAt(0, new Point2d(leftEdgeX + 3, Upper ? yPoint + 6.28 : yPoint - 3), 0, 0, 0);
            groundLasso.AddVertexAt(1, new Point2d(rightEdgeX - 3, Upper ? yPoint + 6.28 : yPoint - 3), -1, 0, 0);
            groundLasso.AddVertexAt(2, groundLasso.GetPoint2dAt(1).Add(new Vector2d(0, -3.28)), 0, 0, 0);
            groundLasso.AddVertexAt(3, groundLasso.GetPoint2dAt(0).Add(new Vector2d(0, -3.28)), -1, 0, 0);
            modSpace.AppendEntity(groundLasso);
            acTrans.AddNewlyCreatedDBObject(groundLasso, true);

            Circle acCirc = new Circle();
            acCirc.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                acCirc.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                acCirc.LinetypeId = lineTypeTable["hidden2"];
            acCirc.LinetypeScale = 5;
            acCirc.Center = groundLasso.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
            acCirc.Radius = 0.37;
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

            Line groundLine1 = new Line();
            groundLine1.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLine1.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLine1.LinetypeId = lineTypeTable["hidden2"];
            groundLine1.LinetypeScale = 5;
            groundLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLine1.StartPoint = groundLasso.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
            groundLine1.EndPoint = new Point3d(lPoint + (rPoint - lPoint) / 2, groundLine1.StartPoint.Y, 0);
            modSpace.AppendEntity(groundLine1);
            acTrans.AddNewlyCreatedDBObject(groundLine1, true);

            Line groundLine2 = new Line();
            groundLine2.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLine2.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLine2.LinetypeId = lineTypeTable["hidden2"];
            groundLine2.LinetypeScale = 5;
            groundLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLine2.StartPoint = groundLine1.EndPoint;
            groundLine2.EndPoint = new Point3d(lPoint + (rPoint - lPoint) / 2, yPoint, 0);
            modSpace.AppendEntity(groundLine2);
            acTrans.AddNewlyCreatedDBObject(groundLine2, true);
        }
    }
}
