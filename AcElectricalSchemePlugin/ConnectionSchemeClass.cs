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
        private static DBText currentTBoxName = null;
        private static Leader currentLeader = null;
        private static Table table = null;

        struct tBoxUnit
        {
            public string Name;
            public int Count;
            public int LastTerminalNumber;

            public tBoxUnit(string name, int count, int lastTerminalNumber)
            {
                Name = name;
                Count = count;
                LastTerminalNumber = lastTerminalNumber;
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
            public List<Terminal> terminals;
            public List<string> equipTerminals;

            public Unit(string _cupboardName, string _tBoxName, string _designation, string _param, string _equipment, string _equipType, string _cableMark, bool _shield, List<Terminal> _terminals, List<string> _equipTerminals)
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
            }
        }

        struct Terminal
        {
            public string boxTerminal1;
            public string boxTerminal2;
            public Terminal(string terminal1, string terminal2)
            {
                boxTerminal1 = terminal1;
                boxTerminal2 = terminal2;
            }
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
                        string[] text = lines[i].Split(';');
                        List<Terminal> terminals = new List<Terminal>();
                        for (int j=8; j<18; j+=2)
                            if (text[j] != "" && text[j+1]!="")
                                terminals.Add(new Terminal(text[j], text[j+1]));
                        List<string> equipTerminals = new List<string>();
                        for (int j = 18; j < 28; j++)
                            if (text[j] != "")
                                equipTerminals.Add(text[j]);
                        bool shield = text[7].ToUpper() == "ДА" || text[7].ToUpper() == "ЕСТЬ" ? true : false;
                        units.Add(new Unit(text[0], text[1], text[2], text[3], text[4], text[5], text[6], shield, terminals, equipTerminals));
                    }
                    firstSheet = true;
                    tBoxes = new List<tBoxUnit>();
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
                        acDoc.Editor.Regen();

                        Point3d startPoint = selection.Value;
                        drawUnits(acTrans, acDb, acModSpace, units, drawSheet(acTrans, acModSpace, acDb, startPoint, firstSheet));
                        acDoc.Editor.Regen();
                        acTrans.Commit();
                        acTrans.Dispose();
                        currentTable = null;
                        currentTBox = null;
                        currentTBoxName = null;
                        currentLeader = null;
                    }
                }
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

            insertSheet(acTrans, modSpace, acdb, new Point3d(shieldPoly.GetPoint2dAt(0).X - 92, shieldPoly.GetPoint2dAt(0).Y + 37, 0), first);

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
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = gndPoly.GetPoint3dAt(0).Add(new Vector3d(0, 2, 0));
            text.TextString = "Шина функционального заземления";
            text.HorizontalMode = TextHorizontalMode.TextLeft;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            table = new Table();
            table.Position = shieldPoly.GetPoint3dAt(0).Add(new Vector3d(-14, -272, 0));
            table.SetSize(4, 1);
            table.TableStyle = acdb.Tablestyle;

            table.SetTextHeight(0, 0, 2.5);
            table.Cells[0, 0].TextString = "Тип оборудования";
            table.SetAlignment(0, 0, CellAlignment.MiddleCenter);
            table.Rows[0].IsMergeAllEnabled = false;

            table.SetTextHeight(1, 0, 2.5);
            table.Cells[1, 0].TextString = "Обозначение по проекту";
            table.SetAlignment(1, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(2, 0, 2.5);
            table.Cells[2, 0].TextString = "Параметры";
            table.SetAlignment(2, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(3, 0, 2.5);
            table.Cells[3, 0].TextString = "Оборудоание";
            table.SetAlignment(3, 0, CellAlignment.MiddleCenter);

            table.Columns[0].Width = 31; 
            table.GenerateLayout();
            modSpace.AppendEntity(table);
            acTrans.AddNewlyCreatedDBObject(table, true);
            currentTable = table;

            return shieldPoly;
        }

        private static void drawUnits(Transaction acTrans, Database acdb, BlockTableRecord modSpace, List<Unit> units, Polyline shield)
        {
            Polyline tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[0].cupboardName);
            Polyline prevPoly = tBoxPoly;
            Polyline prevTermPoly = null;
            string prevTerminal=string.Empty;
            string prevCupboard = units[0].cupboardName;
            int cablecount = 1;
            for (int j = 0; j < units.Count;)
            {
                bool aborted = false;
                double leftEdgeX = 0;
                double rightEdgeX = 0;
                cablecount = 1;
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
                }
                using (Transaction trans = acdb.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < units[j].terminals.Count; i++)
                    {
                        string terminalTag = units[j].terminals[i].boxTerminal1.Split('[', ']')[1];
                        string terminal1 = units[j].terminals[i].boxTerminal1.Split('[', ']')[2];
                        string terminal2 = units[j].terminals[i].boxTerminal2.Split('[', ']')[2];
                        if (prevTerminal == String.Empty)
                        {
                            prevTerminal = terminalTag;
                            DBText text = new DBText();
                            text.SetDatabaseDefaults();
                            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                            text.TextString = terminalTag;
                            text.HorizontalMode = TextHorizontalMode.TextCenter;
                            text.AlignmentPoint = text.Position;
                            modSpace.AppendEntity(text);
                            acTrans.AddNewlyCreatedDBObject(text, true);

                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal1, terminal2, false, out lowestPoint, units[j].designation, 1);
                            leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
                        }
                        else if (prevTerminal == terminalTag)
                        {
                            Point2d p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(56, 0));
                            prevPoly.SetPointAt(1, p1);
                            Point2d p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(56, 0));
                            prevPoly.SetPointAt(2, p2);
                            if (p1.X >= (firstSheet ? shield.GetPoint2dAt(1).X - 170 : shield.GetPoint2dAt(1).X))
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
                                break;
                            }
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevTermPoly, terminal1, terminal2, true, out lowestPoint, units[j].designation, 1);
                            if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
                        }
                        else
                        {
                            tBoxPoly = new Polyline();
                            tBoxPoly.SetDatabaseDefaults();
                            tBoxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            tBoxPoly.Closed = true;
                            tBoxPoly.AddVertexAt(0, prevPoly.GetPoint2dAt(1).Add(new Vector2d(30, 0)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(1, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(27, 0)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(2, tBoxPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
                            tBoxPoly.AddVertexAt(3, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
                            modSpace.AppendEntity(tBoxPoly);
                            acTrans.AddNewlyCreatedDBObject(tBoxPoly, true);
                            prevPoly = tBoxPoly;
                            prevTermPoly = null;
                            prevTerminal = terminalTag;

                            DBText text = new DBText();
                            text.SetDatabaseDefaults();
                            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                            text.TextString = terminalTag;
                            text.HorizontalMode = TextHorizontalMode.TextCenter;
                            text.AlignmentPoint = text.Position;
                            modSpace.AppendEntity(text);
                            acTrans.AddNewlyCreatedDBObject(text, true);
                            if (tBoxPoly.GetPoint2dAt(1).X >= (firstSheet ? shield.GetPoint2dAt(1).X - 170 : shield.GetPoint2dAt(1).X))
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
                                break;
                            }
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal1, terminal2, false, out lowestPoint, units[j].designation, cablecount);
                            cablecount += 2;
                            if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
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
                        double width = drawCable(acTrans, modSpace, new Point3d(leftEdgeX + 3, lowestPoint.Y, 0), new Point3d(rightEdgeX - 3, lowestPoint.Y, 0), units[j]);

                        currentUnit = units[j];

                        currentTable.InsertColumns(currentTable.Columns.Count, width+40, 1);
                        currentTable.UnmergeCells(currentTable.Rows[0]);
                        currentTable.Columns[currentTable.Columns.Count - 1].TextHeight = 2.5;
                        currentTable.Cells[0, currentTable.Columns.Count-1].TextString = units[j].equipType==""?"-":units[j].equipType;
                        currentTable.Cells[1, currentTable.Columns.Count-1].TextString = units[j].designation;
                        currentTable.Cells[2, currentTable.Columns.Count-1].TextString = units[j].param;
                        currentTable.Cells[3, currentTable.Columns.Count-1].TextString = units[j].equipment;
                        currentTable.GenerateLayout();

                        units.RemoveAt(j);
                        trans.Commit();
                    }
                }
            }
        }

        private static Polyline drawShieldTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Polyline shield, string cupbordName)
        {
            Polyline boxPoly = new Polyline();
            boxPoly.SetDatabaseDefaults();
            boxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            boxPoly.Closed = true;
            boxPoly.AddVertexAt(0, shield.GetPoint2dAt(0).Add(new Vector2d(5, -25)), 0, 0, 0);
            boxPoly.AddVertexAt(1, boxPoly.GetPoint2dAt(0).Add(new Vector2d(27, 0)), 0, 0, 0);
            boxPoly.AddVertexAt(2, boxPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            boxPoly.AddVertexAt(3, boxPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(boxPoly);
            acTrans.AddNewlyCreatedDBObject(boxPoly, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
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

        private static Polyline drawTerminal(Transaction acTrans, BlockTableRecord modSpace, Polyline prevPoly, string term1, string term2, bool iteration, out Point3d lowestPoint, string cableMark, int cableNumber)
        {
            Polyline termPoly1 = new Polyline();
            termPoly1.SetDatabaseDefaults();
            termPoly1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly1.Closed = true;
            termPoly1.AddVertexAt(0, prevPoly.GetPoint2dAt(0).Add(new Vector2d(iteration?50:15, 0)), 0, 0, 0);
            termPoly1.AddVertexAt(1, termPoly1.GetPoint2dAt(0).Add(new Vector2d(6, 0)), 0, 0, 0);
            termPoly1.AddVertexAt(2, termPoly1.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            termPoly1.AddVertexAt(3, termPoly1.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(termPoly1);
            acTrans.AddNewlyCreatedDBObject(termPoly1, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = termPoly1.GetPoint3dAt(0).Add(new Vector3d(3, -4, 0));
            text.TextString = term1;
            text.Rotation = 1.5708;
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Polyline termPoly2 = new Polyline();
            termPoly2.SetDatabaseDefaults();
            termPoly2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly2.Closed = true;
            termPoly2.AddVertexAt(0, termPoly1.GetPoint2dAt(1), 0, 0, 0);
            termPoly2.AddVertexAt(1, termPoly2.GetPoint2dAt(0).Add(new Vector2d(6, 0)), 0, 0, 0);
            termPoly2.AddVertexAt(2, termPoly2.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            termPoly2.AddVertexAt(3, termPoly1.GetPoint2dAt(2), 0, 0, 0);
            modSpace.AppendEntity(termPoly2);
            acTrans.AddNewlyCreatedDBObject(termPoly2, true);

            text = new DBText();
            text.SetDatabaseDefaults();
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = termPoly2.GetPoint3dAt(0).Add(new Vector3d(3, -4, 0));
            text.TextString = term2;
            text.Rotation = 1.5708;
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Line cableLineUp1 = new Line();
            cableLineUp1.SetDatabaseDefaults();
            cableLineUp1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp1.StartPoint = new Point3d(termPoly1.GetPoint2dAt(3).X+(termPoly1.GetPoint2dAt(2).X - termPoly1.GetPoint2dAt(3).X) / 2, termPoly1.GetPoint2dAt(3).Y, 0);
            cableLineUp1.EndPoint = cableLineUp1.StartPoint.Add(new Vector3d(0, -44, 0));
            modSpace.AppendEntity(cableLineUp1);
            acTrans.AddNewlyCreatedDBObject(cableLineUp1, true);
           
            MText textLine1 = new MText();
            textLine1.SetDatabaseDefaults();
            textLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            textLine1.Location = cableLineUp1.EndPoint.Add(new Vector3d(-1, 2, 0));
            textLine1.Contents = cableMark + "-" + cableNumber;
            textLine1.Rotation = 1.5708;
            textLine1.Attachment = AttachmentPoint.BottomLeft;
            modSpace.AppendEntity(textLine1);
            acTrans.AddNewlyCreatedDBObject(textLine1, true);

            Line cableLineUp2 = new Line();
            cableLineUp2.SetDatabaseDefaults();
            cableLineUp2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp2.StartPoint = new Point3d(termPoly2.GetPoint2dAt(3).X + (termPoly2.GetPoint2dAt(2).X - termPoly2.GetPoint2dAt(3).X) / 2, termPoly2.GetPoint2dAt(3).Y, 0);
            cableLineUp2.EndPoint = cableLineUp2.StartPoint.Add(new Vector3d(0, -44, 0));
            modSpace.AppendEntity(cableLineUp2);
            acTrans.AddNewlyCreatedDBObject(cableLineUp2, true);

            MText textLine2 = new MText();
            textLine2.SetDatabaseDefaults();
            textLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            textLine2.Location = cableLineUp2.EndPoint.Add(new Vector3d(-1, 2, 0));
            textLine2.Contents = cableMark + "-" + (cableNumber + 1);
            textLine2.Rotation = 1.5708;
            textLine2.Attachment = AttachmentPoint.BottomLeft;
            modSpace.AppendEntity(textLine2);
            acTrans.AddNewlyCreatedDBObject(textLine2, true);

            while (cableLineUp1.Length - 8 < textLine1.ActualWidth || cableLineUp2.Length - 8 < textLine2.ActualWidth)
            {
                cableLineUp1.EndPoint = cableLineUp1.EndPoint.Add(new Vector3d(0, -1, 0));
                textLine1.Location = cableLineUp1.EndPoint.Add(new Vector3d(-1, 2, 0));
                
                cableLineUp2.EndPoint = cableLineUp2.EndPoint.Add(new Vector3d(0, -1, 0));
                textLine2.Location = cableLineUp2.EndPoint.Add(new Vector3d(-1, 2, 0));
            }

            lowestPoint = cableLineUp2.EndPoint;

            return termPoly2;
        }

        private static double drawCable(Transaction acTrans, BlockTableRecord modSpace, Point3d firstCable, Point3d lastCable, Unit unit)
        {
            double leftEdgeX = 0;
            double rightEdgeX = 0;

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
            textName.Location = cableNameCentre;
            textName.Attachment = AttachmentPoint.TopCenter;
            textName.Width = 25;
            textName.Contents = unit.tBoxName == string.Empty ? unit.cupboardName.Split(' ')[1] + "/" + unit.designation : unit.cupboardName.Split(' ')[1] + "/" + unit.tBoxName;
            modSpace.AppendEntity(textName);
            acTrans.AddNewlyCreatedDBObject(textName, true);
            if (textName.Width < textName.ActualWidth)
            {
                textName.Contents = unit.tBoxName == string.Empty ? unit.cupboardName.Split(' ')[1] + " /" + unit.designation : unit.cupboardName.Split(' ')[1] + " /" + unit.tBoxName;
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

            DBText cableMark = new DBText();
            cableMark.SetDatabaseDefaults();
            cableMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableMark.Position = cableLine2.StartPoint.Add(new Vector3d(0, -20, 0));
            cableMark.TextString = unit.cableMark;
            cableMark.Rotation = 1.5708;
            cableMark.VerticalMode = TextVerticalMode.TextBottom;
            cableMark.HorizontalMode = TextHorizontalMode.TextCenter;
            cableMark.AlignmentPoint = cableMark.Position;
            modSpace.AppendEntity(cableMark);
            acTrans.AddNewlyCreatedDBObject(cableMark, true);
            
            Point3d point;
            if (unit.tBoxName != string.Empty)
            {
                point = drawTerminalBox(acTrans, modSpace, cableLine2, unit);
            }
            else
            {
                point = cableLine2.EndPoint;
                currentTBoxName = null;
            }
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = point.X - ((unit.equipTerminals.Count-1) * 5) / 2;
            jumperLineDown.StartPoint = new Point3d(x, point.Y, 0);
            x = x + ((unit.equipTerminals.Count-1) * 5);
            jumperLineDown.EndPoint = new Point3d(x, point.Y, 0);
            modSpace.AppendEntity(jumperLineDown);
            acTrans.AddNewlyCreatedDBObject(jumperLineDown, true);

            for (int i = 0; i < unit.equipTerminals.Count; i++)
            {
                Line cableLineDown = new Line();
                cableLineDown.SetDatabaseDefaults();
                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableLineDown.StartPoint = jumperLineDown.StartPoint.Add(new Vector3d(i * 5, 0, 0));
                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -44, 0));
                modSpace.AppendEntity(cableLineDown);
                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

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

                DBText textLine = new DBText();
                textLine.SetDatabaseDefaults();
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Position = cableLineDown.EndPoint.Add(new Vector3d(0, 20, 0));
                textLine.TextString = unit.designation + "-" + (i+1);
                textLine.Rotation = 1.5708;
                textLine.VerticalMode = TextVerticalMode.TextBottom;
                textLine.HorizontalMode = TextHorizontalMode.TextCenter;
                textLine.AlignmentPoint = textLine.Position;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);
                
                DBText text = new DBText();
                text.SetDatabaseDefaults();
                text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                text.Position = terminal.GetPoint3dAt(0).Add(new Vector3d(2.5, -3.5, 0));
                text.TextString = unit.equipTerminals[i];
                text.HorizontalMode = TextHorizontalMode.TextCenter;
                text.AlignmentPoint = text.Position;
                modSpace.AppendEntity(text);
                acTrans.AddNewlyCreatedDBObject(text, true);

                if (i == 0) leftEdgeX = cableLineDown.StartPoint.X - 3;
                rightEdgeX = cableLineDown.StartPoint.X + 3;
            }
            if (unit.shield)
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
            equip.AddVertexAt(2, equip.GetPoint2dAt(1).Add(new Vector2d(0, -60)), 0, 0, 0);
            equip.AddVertexAt(3, equip.GetPoint2dAt(0).Add(new Vector2d(0, -60)), 0, 0, 0);
            modSpace.AppendEntity(equip);
            acTrans.AddNewlyCreatedDBObject(equip, true);
            
            if (equip.GetPoint2dAt(3).Y < table.Position.Y)
                table.Position = table.Position.Add(new Vector3d(0, - 10, 0));

            return (equip.GetPoint2dAt(1).X - equip.GetPoint2dAt(0).X);
        }

        private static Point3d drawTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Line cableLine, Unit unit)
        {
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = cableLine.EndPoint.X - unit.equipTerminals.Count * 5 / 2;
            jumperLineDown.StartPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            x = x + unit.equipTerminals.Count * 5;
            jumperLineDown.EndPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            modSpace.AppendEntity(jumperLineDown);
            acTrans.AddNewlyCreatedDBObject(jumperLineDown, true);

            double lowestPointY = 0;
            double leftEdgeX = 0;
            double rightEdgeX = 0;
            tBoxUnit tBox;
            int index = 0;
            bool exist = false;
            if (tBoxes.Exists(tbox => tbox.Name == unit.tBoxName))
            {
                index = tBoxes.IndexOf(tBoxes.Find(tbox => tbox.Name == unit.tBoxName));
                tBox = tBoxes[index];
                exist = true;
            }
            else tBox = new tBoxUnit(unit.tBoxName, 0, 1);
            double lowestPoint = 0;
            List<Polyline> polys = new List<Polyline>();
            List<MText> texts = new List<MText>();
            for (int i = 0; i < unit.equipTerminals.Count; i++)
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
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                textLine.Contents = unit.designation + "-" + (i + 1);
                textLine.Rotation = 1.5708;
                textLine.Attachment = AttachmentPoint.BottomLeft;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);
                
                polys.Add(terminal);
                texts.Add(textLine);

                while (textLine.ActualWidth > Math.Abs(terminal.GetPoint2dAt(1).Y) - Math.Abs(terminal.GetPoint2dAt(2).Y) - 7)
                {
                    terminal.SetPointAt(2, terminal.GetPoint2dAt(2).Add(new Vector2d(0, -1)));
                    terminal.SetPointAt(3, terminal.GetPoint2dAt(3).Add(new Vector2d(0, -1)));
                    textLine.Location = terminal.GetPoint3dAt(2).Add(new Vector3d(-1, 6, 0));
                }
                if (i == 0) lowestPoint = terminal.GetPoint2dAt(2).Y;
                else if (lowestPoint > terminal.GetPoint2dAt(2).Y) 
                    lowestPoint = terminal.GetPoint2dAt(2).Y;
            }

            for (int i = 0; i < unit.equipTerminals.Count; i++)
            {
                //Polyline terminal = (Polyline)acTrans.GetObject(polys[i].Id, OpenMode.ForWrite);
                //MText textLine = (MText)acTrans.GetObject(texts[i].Id, OpenMode.ForWrite);
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

            if (unit.shield)
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

                DBText textLine = new DBText();
                textLine.SetDatabaseDefaults();
                textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textLine.Position = terminal.GetPoint3dAt(2).Add(new Vector3d(0, 20, 0));
                textLine.TextString = "шина ф. заземл.";
                textLine.Rotation = 1.5708;
                textLine.VerticalMode = TextVerticalMode.TextBottom;
                textLine.HorizontalMode = TextHorizontalMode.TextCenter;
                textLine.AlignmentPoint = textLine.Position;
                modSpace.AppendEntity(textLine);
                acTrans.AddNewlyCreatedDBObject(textLine, true);

                Line line = new Line();
                line.SetDatabaseDefaults();
                line.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line.StartPoint = terminal.GetPoint3dAt(3).Add(new Vector3d(0, 5, 0));
                line.EndPoint = terminal.GetPoint3dAt(2).Add(new Vector3d(0, 5, 0));
                modSpace.AppendEntity(line);
                acTrans.AddNewlyCreatedDBObject(line, true);

                //DBText text = new DBText();
                //text.SetDatabaseDefaults();
                //text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                //text.Position = line.StartPoint.Add(new Vector3d(2.5, -3.5, 0));
                //text.TextString = unit.boxTerminals[i];
                //text.HorizontalMode = TextHorizontalMode.TextCenter;
                //text.AlignmentPoint = text.Position;
                //modSpace.AppendEntity(text);
                //acTrans.AddNewlyCreatedDBObject(text, true);

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

            if (currentTBoxName == null || currentUnit.tBoxName != unit.tBoxName)
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
                DBText tBoxName = new DBText();
                tBoxName.SetDatabaseDefaults();
                tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxName.Position = tbox.GetPoint3dAt(1).Add(new Vector3d(8, 5, 0));
                tBoxName.TextString = tBox.Name+"."+tBox.Count.ToString();
                tBoxName.VerticalMode = TextVerticalMode.TextBottom;
                tBoxName.HorizontalMode = TextHorizontalMode.TextCenter;
                tBoxName.AlignmentPoint = tBoxName.Position;
                modSpace.AppendEntity(tBoxName);
                acTrans.AddNewlyCreatedDBObject(tBoxName, true);
                currentTBoxName = tBoxName;
                
                Leader acLdr = new Leader();
                acLdr.SetDatabaseDefaults();
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                acLdr.AppendVertex(tbox.GetPoint3dAt(1).Add(new Vector3d(14, 5, 0)));
                acLdr.HasArrowHead = false;
                modSpace.AppendEntity(acLdr);
                acTrans.AddNewlyCreatedDBObject(acLdr, true);
                currentLeader = acLdr;
            }
            else
            {
                currentTBox.SetPointAt(1, new Point2d(jumperLineDown.EndPoint.X + 9, jumperLineDown.StartPoint.Y + 2));
                currentTBox.SetPointAt(2, new Point2d(jumperLineDown.EndPoint.X + 9, jumperOutput.StartPoint.Y - 2));
                if (currentTBox.GetPoint2dAt(3).Y > jumperLineDown.StartPoint.Y)
                    currentTBox.SetPointAt(3, new Point2d(currentTBox.GetPoint2dAt(3).X, jumperOutput.StartPoint.Y - 2));


                currentTBoxName.Erase();
                DBText tBoxName = new DBText();
                tBoxName.SetDatabaseDefaults();
                tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxName.Position = currentTBox.GetPoint3dAt(1).Add(new Vector3d(8, 5, 0));
                tBoxName.TextString = tBox.Name + "." + tBox.Count.ToString();
                tBoxName.VerticalMode = TextVerticalMode.TextBottom;
                tBoxName.HorizontalMode = TextHorizontalMode.TextCenter;
                tBoxName.AlignmentPoint = tBoxName.Position;
                modSpace.AppendEntity(tBoxName);
                acTrans.AddNewlyCreatedDBObject(tBoxName, true);
                currentTBoxName = tBoxName;

                currentLeader.Erase();
                Leader acLdr = new Leader();
                acLdr.SetDatabaseDefaults();
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                acLdr.AppendVertex(currentTBox.GetPoint3dAt(1).Add(new Vector3d(14, 5, 0)));
                acLdr.HasArrowHead = false;
                modSpace.AppendEntity(acLdr);
                acTrans.AddNewlyCreatedDBObject(acLdr, true);
                currentLeader = acLdr;
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
