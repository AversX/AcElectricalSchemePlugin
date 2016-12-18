using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;
using System.Data;
using System.Data.OleDb;

namespace AcElectricalSchemePlugin
{
    static class ConnectionScheme
    {
        private static Document acDoc;
        private static List<Unit> units;
        private static Editor editor;
        private static int curTableNum;
        private static LinetypeTable lineTypeTable;
        private static TextStyleTable tst;
        private static double width = 0;
        private static List<Table> tables = new List<Table>();
        private static int currentSheetNumber = 1;
        private static List<tBox> tboxes;
        private static int curTermNumber;
        private static int curPairNumber;
        private static MText curPairMark;

        struct tBox
        {
            public string Name;
            public int Count;
            public int LastTerminalNumber;
            public int LastShieldNumber;
            public int LastPairNumber;
            public MText textName;
            public Leader ldr;
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
            public List<string> colors;
            public List<string> equipTerminals;
            public List<Point3d> cableOutput;
            public List<Line> cables;
            public MText cableName;
            public bool newSheet; 

            public Unit(string _cupboardName, string _tBoxName, string _designation, string _param, string _equipment, string _equipType, string _cableMark, bool _shield, List<string> _terminals, List<string> _colors, List<string> _equipTerminals)
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
                colors = _colors;
                equipTerminals = _equipTerminals;
                cableOutput = new List<Point3d>();
                cables = new List<Line>();
                cableName = null;
                newSheet = false;
            }
        }

        private static List<Unit> loadData(string path)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO;'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.FillSchema(dataSet, SchemaType.Mapped);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 2; row < dataSet.Tables[0].Rows.Count; row++)
            {
                Unit unit = new Unit();
                unit.cupboardName = dataSet.Tables[0].Rows[row][0].ToString();
                unit.tBoxName = dataSet.Tables[0].Rows[row][1].ToString();
                unit.designation = dataSet.Tables[0].Rows[row][2].ToString();
                unit.param = dataSet.Tables[0].Rows[row][3].ToString();
                unit.equipment = dataSet.Tables[0].Rows[row][4].ToString();
                unit.equipType = dataSet.Tables[0].Rows[row][5].ToString();
                unit.cableMark = dataSet.Tables[0].Rows[row][6].ToString();
                if (dataSet.Tables[0].Rows[row][7].ToString().ToUpper().Contains("ДА")) unit.shield = true;
                else unit.shield = false;
                List<string> terminals = new List<string>();
                for (int i = 8; i < 18; i++)
                    if (dataSet.Tables[0].Rows[row][i].ToString() != "")
                        terminals.Add(dataSet.Tables[0].Rows[row][i].ToString());
                unit.terminals = terminals;
                List<string> colors = new List<string>(dataSet.Tables[0].Rows[row][18].ToString().Split('-'));
                unit.colors = colors;
                List<string> equipTerminals = new List<string>();
                for (int i = 19; i < 29; i++)
                    if (dataSet.Tables[0].Rows[row][i].ToString() != "")
                        equipTerminals.Add(dataSet.Tables[0].Rows[row][i].ToString());
                unit.equipTerminals = equipTerminals;
                units.Add(unit);
            }
            return units;
        }

        public static void DrawScheme()
        {
            units = new List<Unit>();
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "(*.xlsx)|*.xlsx";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                curTableNum = 0;
                width = 0;
                tables = new List<Table>();
                currentSheetNumber = 1;
                tboxes = new List<tBox>();
                curTermNumber = 1;
                curPairNumber = 1;
                currentSheetNumber = 1;
                units = loadData(file.FileName);
                connectionScheme();
            }
        }

        private static void connectionScheme()
        {
            Database acDb = acDoc.Database;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    PromptPointResult selectedPoint = acDoc.Editor.GetPoint("Выберите точку");
                    if (selectedPoint.Status == PromptStatus.OK)
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
                        tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForWrite);
                        loadFonts(acTrans, acModSpace, acDb);
                       
                        Point3d startPoint = selectedPoint.Value;
                        Polyline sh = drawSheet(acTrans, acModSpace, acDb, startPoint);
                        drawUnits(acTrans, acDb, acModSpace, units, sh);
                        drawCables(acTrans, acModSpace, units, sh);
                        for (int i = 0; i < tables.Count; i++)
                        {
                            if (!acBlkTbl.Has(tables[i].Id))
                            {
                                acModSpace.AppendEntity(tables[i]);
                                acTrans.AddNewlyCreatedDBObject(tables[i], true);
                            }
                        }
                        acDoc.Editor.Regen();
                        acTrans.Commit();
                    }
                }
                acDb.Audit(true, true);
            }
        }

        private static Polyline drawSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d prevSheet)
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

            insertSheet(acTrans, modSpace, acdb, new Point3d(shieldPoly.GetPoint2dAt(0).X - 92, shieldPoly.GetPoint2dAt(0).Y + 10, 0));

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
            if (tst.Has("spds 2.5-0.85"))
                text.TextStyleId = tst["spds 2.5-0.85"];
            text.Height = 3;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = gndPoly.GetPoint3dAt(0).Add(new Vector3d(0, 2, 0));
            text.TextString = "Шина функционального заземления";
            text.HorizontalMode = TextHorizontalMode.TextLeft;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Table table = new Table();
            table.Position = shieldPoly.GetPoint3dAt(0).Add(new Vector3d(-30, -325, 0));
            table.SetSize(4, 1);
            table.TableStyle = acdb.Tablestyle;

            table.SetTextHeight(0, 0, 2.5);
            table.Cells[0, 0].TextString = "Тип оборудования";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[0, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(0, 0, CellAlignment.MiddleCenter);
            table.Rows[0].IsMergeAllEnabled = false;

            table.SetTextHeight(1, 0, 2.5);
            table.Cells[1, 0].TextString = "Обозначение по проекту";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[1, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(1, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(2, 0, 2.5);
            table.Cells[2, 0].TextString = "Параметры";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[2, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(2, 0, CellAlignment.MiddleCenter);

            table.SetTextHeight(3, 0, 2.5);
            table.Cells[3, 0].TextString = "Оборудоание";
            if (tst.Has("spds 2.5-0.85"))
                table.Cells[3, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.SetAlignment(3, 0, CellAlignment.MiddleCenter);

            table.Columns[0].Width = 30;
            table.GenerateLayout();
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
            int color = 0;
            for (int j = 0; j < units.Count; j++)
            {
                Unit un = units[j];
                un.cableOutput = new List<Point3d>();
                units[j] = un;
                int offset = 15;
                color = 0;
                bool aborted = false;
                double leftEdgeX = 0;
                Point3d lowestPoint = Point3d.Origin;
                double l=0;
                double r=0;
                bool shielded = false;
                if (prevCupboard != units[j].cupboardName)
                {
                    shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
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
                            text.Height = 3;
                            if (tst.Has("spds 2.5-0.85"))
                                text.TextStyleId = tst["spds 2.5-0.85"];
                            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                            text.TextString = terminalTag;
                            text.HorizontalMode = TextHorizontalMode.TextCenter;
                            text.AlignmentPoint = text.Position;
                            modSpace.AppendEntity(text);
                            acTrans.AddNewlyCreatedDBObject(text, true);
                            newJ = false;
                            shielded = false;

                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                            if (color==0) l = prevTermPoly.GetPoint2dAt(0).X;
                            if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                            else 
                            { 
                                color = 0; 
                                offset = 15; 
                                r = prevTermPoly.GetPoint2dAt(1).X; 
                                units[j].cableOutput.Add(drawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield));
                                shielded = true; 
                            }
                            leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                        }
                        else if (prevTerminal == terminalTag)
                        {
                            Point2d p1;
                            Point2d p2;
                            if (newJ)
                            {
                                p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(offset+60, 0));
                                prevPoly.SetPointAt(1, p1);
                                p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(offset+60, 0));
                                prevPoly.SetPointAt(2, p2);
                            }
                            else
                            {
                                p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(offset+6, 0));
                                prevPoly.SetPointAt(1, p1);
                                p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(offset+6, 0));
                                prevPoly.SetPointAt(2, p2);
                            }
                            if (p1.X >= shield.GetPoint2dAt(1).X - 170)
                            {
                                trans.Abort();
                                aborted = true;
                                shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950,0,0)));
                                tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
                                prevPoly = tBoxPoly;
                                prevTermPoly = null;
                                prevTerminal = string.Empty;
                                newJ = false;
                                Unit unit = units[j];
                                unit.newSheet = true;
                                units[j] = unit;
                                break;
                            }
                            shielded = false;
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                            if (color==0) l = prevTermPoly.GetPoint2dAt(0).X;
                            if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                            else 
                            { 
                                color = 0; 
                                offset = 15; 
                                r = prevTermPoly.GetPoint2dAt(1).X;
                                units[j].cableOutput.Add(drawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield));
                                shielded = true; 
                            }
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
                            text.Height = 3;
                            if (tst.Has("spds 2.5-0.85"))
                                text.TextStyleId = tst["spds 2.5-0.85"];
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
                                aborted = true;
                                shield = drawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                tBoxPoly = drawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
                                prevPoly = tBoxPoly;
                                prevTermPoly = null;
                                prevTerminal = string.Empty;
                                newJ = false;
                                Unit unit = units[j];
                                unit.newSheet = true;
                                units[j] = unit;
                                break;
                            }
                            shielded = false;
                            offset = 15;
                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                            if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                            if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                            else { 
                                color = 0; 
                                offset = 15; 
                                r = prevTermPoly.GetPoint2dAt(1).X;
                                units[j].cableOutput.Add(drawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield));
                                shielded = true; 
                            }
                            if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                        }
                    }
                    if (!shielded && !aborted)
                    {
                        r = prevTermPoly.GetPoint2dAt(1).X;
                        units[j].cableOutput.Add(drawGnd(acTrans, modSpace, l-3, r+3, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield));
                    }
                    if (!aborted) trans.Commit(); //units[j] = drawCable(acTrans, modSpace, new Point3d(leftEdgeX + 3, lowestPoint.Y, 0), new Point3d(rightEdgeX - 3, lowestPoint.Y, 0), units[j]);
                    else j--;
                }
                newJ = true;
            }
        }

        private static Point3d drawGnd(Transaction acTrans, BlockTableRecord modSpace, double leftEdgeX, double rightEdgeX, double Y, Point3d lowestPoint, Polyline shield)
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
            groundLasso.AddVertexAt(0, new Point2d(leftEdgeX + 3, Y - 3), 0, 0, 0);
            groundLasso.AddVertexAt(1, new Point2d(rightEdgeX - 3, Y - 3), -1, 0, 0);
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
            groundLine1.EndPoint = new Point3d(rightEdgeX + 4.64, groundLine1.StartPoint.Y, 0);
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

            Line cableLine = new Line();
            cableLine.SetDatabaseDefaults();
            cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLine.StartPoint = new Point3d(leftEdgeX + 3, lowestPoint.Y, 0);
            cableLine.EndPoint = new Point3d(rightEdgeX - 3, lowestPoint.Y, 0);
            modSpace.AppendEntity(cableLine);
            acTrans.AddNewlyCreatedDBObject(cableLine, true);

            Line cableLineDown = new Line();
            cableLineDown.SetDatabaseDefaults();
            cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineDown.StartPoint = new Point3d(leftEdgeX + (rightEdgeX - leftEdgeX) / 2, lowestPoint.Y, 0);
            cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(cableLineDown);
            acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

            return cableLineDown.EndPoint;
        }

        private static void drawGnd(Transaction acTrans, BlockTableRecord modSpace, double leftEdgeX, double rightEdgeX, double Y, double rightestPoint, Polyline shield)
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
            groundLasso.AddVertexAt(0, new Point2d(leftEdgeX + 3, Y - 3), 0, 0, 0);
            groundLasso.AddVertexAt(1, new Point2d(rightEdgeX - 3, Y - 3), -1, 0, 0);
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
            groundLine1.EndPoint = new Point3d(rightestPoint, groundLine1.StartPoint.Y, 0);
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

        private static Point3d drawTBoxGnd(Transaction acTrans, BlockTableRecord modSpace, Point3d lastTermPoint, double leftEdgeX, double rightEdgeX, double upY, double downY)
        {
            Polyline termPoly = new Polyline();
            termPoly.SetDatabaseDefaults();
            termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly.Closed = true;
            termPoly.AddVertexAt(0, new Point2d(lastTermPoint.X, lastTermPoint.Y), 0, 0, 0);
            termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
            termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(1).Add(new Vector2d(0, -10)), 0, 0, 0);
            termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(0).Add(new Vector2d(0, -10)), 0, 0, 0);
            modSpace.AppendEntity(termPoly);
            acTrans.AddNewlyCreatedDBObject(termPoly, true);

            MText termText = new MText();
            termText.SetDatabaseDefaults();
            termText.TextHeight = 2.5;
            if (tst.Has("spds 2.5-0.85"))
                termText.TextStyleId = tst["spds 2.5-0.85"];
            termText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2;
            double y = termPoly.GetPoint3dAt(2).Y + (termPoly.GetPoint3dAt(1).Y - termPoly.GetPoint3dAt(2).Y) / 2;
            termText.Location = new Point3d(x, y, 0);
            if (upY != downY) termText.Contents = "";
            else termText.Contents = "OSH";
            termText.Rotation = 1.5708;
            termText.Attachment = AttachmentPoint.MiddleCenter;
            modSpace.AppendEntity(termText);
            acTrans.AddNewlyCreatedDBObject(termText, true);

            #region Up
            Polyline groundLassoUp = new Polyline();
            groundLassoUp.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLassoUp.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLassoUp.LinetypeId = lineTypeTable["hidden2"];
            groundLassoUp.LinetypeScale = 5;
            groundLassoUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLassoUp.Closed = true;
            groundLassoUp.AddVertexAt(0, new Point2d(leftEdgeX, upY - 3), 0, 0, 0);
            groundLassoUp.AddVertexAt(1, new Point2d(rightEdgeX, upY - 3), -1, 0, 0);
            groundLassoUp.AddVertexAt(2, groundLassoUp.GetPoint2dAt(1).Add(new Vector2d(0, -3.28)), 0, 0, 0);
            groundLassoUp.AddVertexAt(3, groundLassoUp.GetPoint2dAt(0).Add(new Vector2d(0, -3.28)), -1, 0, 0);
            modSpace.AppendEntity(groundLassoUp);
            acTrans.AddNewlyCreatedDBObject(groundLassoUp, true);

            Line groundLine1Up = new Line();
            groundLine1Up.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLine1Up.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLine1Up.LinetypeId = lineTypeTable["hidden2"];
            groundLine1Up.LinetypeScale = 5;
            groundLine1Up.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLine1Up.StartPoint = groundLassoUp.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
            groundLine1Up.EndPoint = new Point3d(termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2, groundLine1Up.StartPoint.Y, 0);
            modSpace.AppendEntity(groundLine1Up);
            acTrans.AddNewlyCreatedDBObject(groundLine1Up, true);

            Circle acCircUp = new Circle();
            acCircUp.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                acCircUp.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                acCircUp.LinetypeId = lineTypeTable["hidden2"];
            acCircUp.LinetypeScale = 5;
            acCircUp.Center = groundLassoUp.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
            acCircUp.Radius = 0.37;
            modSpace.AppendEntity(acCircUp);
            acTrans.AddNewlyCreatedDBObject(acCircUp, true);

            ObjectIdCollection acObjIdCollUp = new ObjectIdCollection();
            acObjIdCollUp.Add(acCircUp.ObjectId);

            Hatch acHatchUp = new Hatch();
            modSpace.AppendEntity(acHatchUp);
            acTrans.AddNewlyCreatedDBObject(acHatchUp, true);
            acHatchUp.SetDatabaseDefaults();
            acHatchUp.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            acHatchUp.Associative = true;
            acHatchUp.AppendLoop(HatchLoopTypes.Outermost, acObjIdCollUp);
            acHatchUp.EvaluateHatch(true);

            Line groundLine2Up = new Line();
            groundLine2Up.SetDatabaseDefaults();
            if (lineTypeTable.Has("штриховая2"))
                groundLine2Up.LinetypeId = lineTypeTable["штриховая2"];
            else if (lineTypeTable.Has("hidden2"))
                groundLine2Up.LinetypeId = lineTypeTable["hidden2"];
            groundLine2Up.LinetypeScale = 5;
            groundLine2Up.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            groundLine2Up.StartPoint = groundLine1Up.EndPoint;
            groundLine2Up.EndPoint = termPoly.GetPoint3dAt(0).Add(new Vector3d((termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X)/2, 0, 0));
            modSpace.AppendEntity(groundLine2Up);
            acTrans.AddNewlyCreatedDBObject(groundLine2Up, true);
            #endregion

            if (upY != downY)
            {
                #region Down
                Polyline groundLassoDown = new Polyline();
                groundLassoDown.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    groundLassoDown.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    groundLassoDown.LinetypeId = lineTypeTable["hidden2"];
                groundLassoDown.LinetypeScale = 5;
                groundLassoDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLassoDown.Closed = true;
                groundLassoDown.AddVertexAt(0, new Point2d(leftEdgeX, downY + 6), 0, 0, 0);
                groundLassoDown.AddVertexAt(1, new Point2d(rightEdgeX, downY + 6), -1, 0, 0);
                groundLassoDown.AddVertexAt(2, groundLassoDown.GetPoint2dAt(1).Add(new Vector2d(0, -3.28)), 0, 0, 0);
                groundLassoDown.AddVertexAt(3, groundLassoDown.GetPoint2dAt(0).Add(new Vector2d(0, -3.28)), -1, 0, 0);
                modSpace.AppendEntity(groundLassoDown);
                acTrans.AddNewlyCreatedDBObject(groundLassoDown, true);

                Line groundLine1Down = new Line();
                groundLine1Down.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    groundLine1Down.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    groundLine1Down.LinetypeId = lineTypeTable["hidden2"];
                groundLine1Down.LinetypeScale = 5;
                groundLine1Down.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLine1Down.StartPoint = groundLassoDown.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
                groundLine1Down.EndPoint = new Point3d(termPoly.GetPoint3dAt(3).X + (termPoly.GetPoint3dAt(2).X - termPoly.GetPoint3dAt(3).X) / 2, groundLine1Down.StartPoint.Y, 0);
                modSpace.AppendEntity(groundLine1Down);
                acTrans.AddNewlyCreatedDBObject(groundLine1Down, true);

                Circle acCircDown = new Circle();
                acCircDown.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    acCircDown.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    acCircDown.LinetypeId = lineTypeTable["hidden2"];
                acCircDown.LinetypeScale = 5;
                acCircDown.Center = groundLassoDown.GetPoint3dAt(1).Add(new Vector3d(1.64, -1.64, 0));
                acCircDown.Radius = 0.37;
                modSpace.AppendEntity(acCircDown);
                acTrans.AddNewlyCreatedDBObject(acCircDown, true);

                ObjectIdCollection acObjIdCollDown = new ObjectIdCollection();
                acObjIdCollDown.Add(acCircDown.ObjectId);

                Hatch acHatchDown = new Hatch();
                modSpace.AppendEntity(acHatchDown);
                acTrans.AddNewlyCreatedDBObject(acHatchDown, true);
                acHatchDown.SetDatabaseDefaults();
                acHatchDown.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                acHatchDown.Associative = true;
                acHatchDown.AppendLoop(HatchLoopTypes.Outermost, acObjIdCollDown);
                acHatchDown.EvaluateHatch(true);

                Line groundLine2Down = new Line();
                groundLine2Down.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    groundLine2Down.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    groundLine2Down.LinetypeId = lineTypeTable["hidden2"];
                groundLine2Down.LinetypeScale = 5;
                groundLine2Down.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLine2Down.StartPoint = groundLine1Down.EndPoint;
                groundLine2Down.EndPoint = termPoly.GetPoint3dAt(3).Add(new Vector3d((termPoly.GetPoint3dAt(2).X - termPoly.GetPoint3dAt(3).X) / 2, 0, 0));
                modSpace.AppendEntity(groundLine2Down);
                acTrans.AddNewlyCreatedDBObject(groundLine2Down, true);
                #endregion
            }

            return termPoly.GetPoint3dAt(1);
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
            if (tst.Has("spds 2.5-0.85"))
                text.TextStyleId = tst["spds 2.5-0.85"];
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

        private static Polyline drawTerminal(Transaction acTrans, BlockTableRecord modSpace, Point2d point, string terminal, out Point3d lowestPoint, Unit unit, int color, int cableNumber)
        {
            Polyline termPoly = new Polyline();
            termPoly.SetDatabaseDefaults();
            termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly.Closed = true;
            termPoly.AddVertexAt(0, point, 0, 0, 0);
            termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(6, 0)), 0, 0, 0);
            termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(termPoly);
            acTrans.AddNewlyCreatedDBObject(termPoly, true);

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.Height = 3;
            if (tst.Has("spds 2.5-0.85"))
                text.TextStyleId = tst["spds 2.5-0.85"];
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
            cableLineUp.EndPoint = cableLineUp.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(cableLineUp);
            acTrans.AddNewlyCreatedDBObject(cableLineUp, true);

            Polyline cablePoly = new Polyline();
            cablePoly.SetDatabaseDefaults();
            cablePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cablePoly.Closed = true;
            cablePoly.AddVertexAt(0, new Point2d(cableLineUp.EndPoint.X-2.5, cableLineUp.EndPoint.Y), 0, 0, 0);
            cablePoly.AddVertexAt(1, cablePoly.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
            cablePoly.AddVertexAt(2, cablePoly.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
            cablePoly.AddVertexAt(3, cablePoly.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
            modSpace.AppendEntity(cablePoly);
            acTrans.AddNewlyCreatedDBObject(cablePoly, true);

            MText textLine = new MText();
            textLine.SetDatabaseDefaults();
            textLine.TextHeight = 2.5;
            if (tst.Has("spds 2.5-0.85"))
                textLine.TextStyleId = tst["spds 2.5-0.85"];
            textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = cablePoly.GetPoint3dAt(0).X + (cablePoly.GetPoint3dAt(1).X - cablePoly.GetPoint3dAt(0).X)/2;
            double y = cablePoly.GetPoint3dAt(2).Y + (cablePoly.GetPoint3dAt(1).Y - cablePoly.GetPoint3dAt(2).Y)/2;
            textLine.Location = new Point3d(x, y, 0);
            textLine.Contents = unit.designation + "/" + cableNumber;
            textLine.Rotation = 1.5708;
            textLine.Attachment = AttachmentPoint.MiddleCenter;
            modSpace.AppendEntity(textLine);
            acTrans.AddNewlyCreatedDBObject(textLine, true);

            Line cableLineDown = new Line();
            cableLineDown.SetDatabaseDefaults();
            cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineDown.StartPoint = new Point3d(x, cablePoly.GetPoint2dAt(3).Y, 0);
            cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -7, 0));
            modSpace.AppendEntity(cableLineDown);
            acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

            MText colorMark = new MText();
            colorMark.SetDatabaseDefaults();
            colorMark.TextHeight = 2.5;
            if (tst.Has("spds 2.5-0.85"))
                colorMark.TextStyleId = tst["spds 2.5-0.85"];
            colorMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            x = cableLineDown.EndPoint.X - 1;
            y = cableLineDown.EndPoint.Y + (cableLineDown.StartPoint.Y - cableLineDown.EndPoint.Y) / 2;
            colorMark.Location = new Point3d(x, y, 0);
            colorMark.Contents = unit.colors[color];
            colorMark.Rotation = 1.5708;
            colorMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(colorMark);
            acTrans.AddNewlyCreatedDBObject(colorMark, true);
            
            lowestPoint = cableLineDown.EndPoint;
            return termPoly;
        }

        private static void drawCables(Transaction acTrans, BlockTableRecord modSpace, List<Unit> units, Polyline shield)
        {
            Line cableLineUp = new Line();
            cableLineUp.SetDatabaseDefaults();
            cableLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp.StartPoint = units[0].cableOutput[0];
            cableLineUp.EndPoint = cableLineUp.StartPoint;
            modSpace.AppendEntity(cableLineUp);
            acTrans.AddNewlyCreatedDBObject(cableLineUp, true);

            List<Point3d> points = new List<Point3d>();
            List<int> index = new List<int>();

            string prevTbox = units[0].tBoxName;
            //tBox tbox = new tBox();
            //tbox.Name = units[0].tBoxName;

            for (int i = 0; i < units.Count; i++)
            {
                if (units[i].tBoxName == prevTbox && !units[i].newSheet)
                    for (int j = 0; j < units[i].cableOutput.Count; j++)
                    {
                        cableLineUp.EndPoint = units[i].cableOutput[j];
                        points.Add(units[i].cableOutput[j]);
                        index.Add(i);
                    }
                else
                {
                    Line cableLine = new Line();
                    cableLine.SetDatabaseDefaults();
                    cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cableLine.StartPoint = new Point3d(cableLineUp.StartPoint.X + (cableLineUp.EndPoint.X - cableLineUp.StartPoint.X) / 2, cableLineUp.EndPoint.Y, 0);
                    cableLine.EndPoint = cableLine.StartPoint.Add(new Vector3d(0, -70, 0));
                    modSpace.AppendEntity(cableLine);
                    acTrans.AddNewlyCreatedDBObject(cableLine, true);

                    MText cableMark = new MText();
                    cableMark.SetDatabaseDefaults();
                    cableMark.TextHeight = 3;
                    if (tst.Has("spds 2.5-0.85"))
                        cableMark.TextStyleId = tst["spds 2.5-0.85"];
                    cableMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cableMark.Location = cableLine.EndPoint.Add(new Vector3d(-1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                    cableMark.Contents = units[i-1].cableMark;
                    cableMark.Rotation = 1.5708;
                    cableMark.Attachment = AttachmentPoint.BottomCenter;
                    modSpace.AppendEntity(cableMark);
                    acTrans.AddNewlyCreatedDBObject(cableMark, true);

                    MText textName = new MText();
                    textName.SetDatabaseDefaults();
                    textName.TextHeight = 3;
                    if (tst.Has("spds 2.5-0.85"))
                        textName.TextStyleId = tst["spds 2.5-0.85"];
                    textName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    textName.Location = cableLine.EndPoint.Add(new Vector3d(1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                    textName.Attachment = AttachmentPoint.TopCenter;
                    textName.Rotation = 1.5708;
                    textName.Contents = units[i - 1].tBoxName == string.Empty ? units[i - 1].cupboardName.Split(' ')[1] + "/" + units[i - 1].designation : units[i - 1].cupboardName.Split(' ')[1] + "/" + units[i - 1].tBoxName;
                    modSpace.AppendEntity(textName);
                    acTrans.AddNewlyCreatedDBObject(textName, true);

                    if (units[i - 1].shield)
                        drawGnd(acTrans, modSpace, cableLine.StartPoint.X - 6, cableLine.StartPoint.X + 6, cableLine.StartPoint.Y, cableLineUp.EndPoint.X + 12.27, shield);
                    
                    List<Point3d> equipPoints = new List<Point3d>();
                    if (units[i - 1].tBoxName != string.Empty)
                    {
                        equipPoints = new List<Point3d>();

                        Line cableLineDown = new Line();
                        cableLineDown.SetDatabaseDefaults();
                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        cableLineDown.StartPoint = new Point3d(cableLineUp.StartPoint.X, cableLine.EndPoint.Y, 0);
                        cableLineDown.EndPoint = new Point3d(cableLineUp.EndPoint.X, cableLine.EndPoint.Y, 0);
                        modSpace.AppendEntity(cableLineDown);
                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                        Point3d lastTermPoint = new Point3d();
                        int prevIndex = 0;
                        curTermNumber = 1;
                        curPairNumber = 1;
                        Point3d point = drawTerminalBoxUnit(acTrans, modSpace, index[0], points[0], cableLineDown.EndPoint, out lastTermPoint);

                        Line tBoxJumper = new Line();
                        tBoxJumper.SetDatabaseDefaults();
                        tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxJumper.StartPoint = point;
                        tBoxJumper.EndPoint = point;
                        modSpace.AppendEntity(tBoxJumper);
                        acTrans.AddNewlyCreatedDBObject(tBoxJumper, true);

                        for (int k = 1; k < points.Count; k++)
                        {
                            if (index[prevIndex] == index[k])
                            {
                                curPairNumber++;
                                point = drawTerminalBoxUnit(acTrans, modSpace, index[k], points[k], cableLineDown.EndPoint, out lastTermPoint);
                                tBoxJumper.EndPoint = point;
                            }
                            else
                            {
                                Line tBoxLineOut = new Line();
                                tBoxLineOut.SetDatabaseDefaults();
                                tBoxLineOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                tBoxLineOut.StartPoint = tBoxJumper.StartPoint + (tBoxJumper.EndPoint - tBoxJumper.StartPoint)/2;
                                tBoxLineOut.EndPoint = tBoxLineOut.StartPoint.Add(new Vector3d(0, -15, 0));
                                modSpace.AppendEntity(tBoxLineOut);
                                acTrans.AddNewlyCreatedDBObject(tBoxLineOut, true);
                                
                                if (curPairNumber==1)
                                    curPairMark.Erase();

                                equipPoints.Add(tBoxLineOut.EndPoint);

                                prevIndex = k;

                                curPairNumber = 1;
                                curTermNumber = 1;

                                point = drawTerminalBoxUnit(acTrans, modSpace, index[k], points[k], cableLineDown.EndPoint, out lastTermPoint);
                                
                                tBoxJumper = new Line();
                                tBoxJumper.SetDatabaseDefaults();
                                tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                tBoxJumper.StartPoint = point;
                                tBoxJumper.EndPoint = point;
                                modSpace.AppendEntity(tBoxJumper);
                                acTrans.AddNewlyCreatedDBObject(tBoxJumper, true);
                            }
                        }

                        Polyline tBoxFrame = new Polyline();
                        tBoxFrame.SetDatabaseDefaults();
                        if (lineTypeTable.Has("штриховая2"))
                            tBoxFrame.LinetypeId = lineTypeTable["штриховая2"];
                        else if (lineTypeTable.Has("hidden2"))
                            tBoxFrame.LinetypeId = lineTypeTable["hidden2"];
                        tBoxFrame.LinetypeScale = 5;
                        tBoxFrame.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxFrame.Closed = true;
                        tBoxFrame.AddVertexAt(0, new Point2d(cableLineDown.StartPoint.X - 12, cableLineDown.EndPoint.Y + 7), 0, 0, 0);
                        tBoxFrame.AddVertexAt(1, new Point2d(cableLineDown.EndPoint.X + 14, cableLineDown.EndPoint.Y + 7), 0, 0, 0);
                        tBoxFrame.AddVertexAt(2, new Point2d(cableLineDown.EndPoint.X + 14, tBoxJumper.EndPoint.Y-2), 0, 0, 0);
                        tBoxFrame.AddVertexAt(3, new Point2d(cableLineDown.StartPoint.X - 12, tBoxJumper.EndPoint.Y-2), 0, 0, 0);
                        modSpace.AppendEntity(tBoxFrame);
                        acTrans.AddNewlyCreatedDBObject(tBoxFrame, true);

                        MText tBoxName = new MText();
                        tBoxName.SetDatabaseDefaults();
                        tBoxName.TextHeight = 3;
                        if (tst.Has("spds 2.5-0.85"))
                            tBoxName.TextStyleId = tst["spds 2.5-0.85"];
                        tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxName.Location = tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(1, 6, 0));
                        tBoxName.Contents = prevTbox + ".";// +tBox.Count.ToString();
                        tBoxName.Attachment = AttachmentPoint.BottomLeft;
                        modSpace.AppendEntity(tBoxName);
                        acTrans.AddNewlyCreatedDBObject(tBoxName, true);

                        Leader acLdr = new Leader();
                        acLdr.SetDatabaseDefaults();
                        acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                        acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                        acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(tBoxName.ActualWidth + 1, 5, 0)));
                        acLdr.HasArrowHead = false;
                        modSpace.AppendEntity(acLdr);
                        acTrans.AddNewlyCreatedDBObject(acLdr, true);

                        Line tBoxLineOutLast = new Line();
                        tBoxLineOutLast.SetDatabaseDefaults();
                        tBoxLineOutLast.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxLineOutLast.StartPoint = tBoxJumper.StartPoint + (tBoxJumper.EndPoint - tBoxJumper.StartPoint) / 2;
                        tBoxLineOutLast.EndPoint = tBoxLineOutLast.StartPoint.Add(new Vector3d(0, -15, 0));
                        modSpace.AppendEntity(tBoxLineOutLast);
                        acTrans.AddNewlyCreatedDBObject(tBoxLineOutLast, true);
                        equipPoints.Add(tBoxLineOutLast.EndPoint);
                        if (curPairNumber == 1)
                            curPairMark.Erase();

                        if (units[i - 1].shield)
                            drawTBoxGnd(acTrans, modSpace, lastTermPoint, cableLine.StartPoint.X - 3, cableLine.StartPoint.X + 3, cableLine.EndPoint.Y+8, cableLine.EndPoint.Y+8);
                    }
                    else
                    {
                        cableLine.EndPoint = cableLine.EndPoint.Add(new Vector3d(0, -135, 0));
                        cableMark.Location = cableLine.EndPoint.Add(new Vector3d(-1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                        textName.Location = cableLine.EndPoint.Add(new Vector3d(1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                        equipPoints.Add(cableLine.EndPoint);
                    }                        
                    for (int k = 0; k < equipPoints.Count; k++)
                    {
                        double length = (units[index[k]].equipTerminals.Count - 2) * 5 + 5;
                        Line equipJumper = new Line();
                        equipJumper.SetDatabaseDefaults();
                        equipJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        equipJumper.StartPoint = equipPoints[k].Add(new Vector3d(-length / 2, 0, 0));
                        equipJumper.EndPoint = equipJumper.StartPoint.Add(new Vector3d(length, 0, 0));
                        modSpace.AppendEntity(equipJumper);
                        acTrans.AddNewlyCreatedDBObject(equipJumper, true);

                        double lowestPoint = 0;
                        int color = 0;
                        for (int c = 0; c < units[index[k]].equipTerminals.Count; c++)
                        {
                            Line equipLine = new Line();
                            equipLine.SetDatabaseDefaults();
                            equipLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            equipLine.StartPoint = equipJumper.StartPoint.Add(new Vector3d(5 * c, 0, 0));
                            equipLine.EndPoint = equipLine.StartPoint.Add(new Vector3d(0, -7, 0));
                            modSpace.AppendEntity(equipLine);
                            acTrans.AddNewlyCreatedDBObject(equipLine, true);

                            MText colorMark = new MText();
                            colorMark.SetDatabaseDefaults();
                            colorMark.TextHeight = 2.5;
                            if (tst.Has("spds 2.5-0.85"))
                                colorMark.TextStyleId = tst["spds 2.5-0.85"];
                            colorMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            double x = equipLine.EndPoint.X - 1;
                            double y = equipLine.EndPoint.Y + (equipLine.StartPoint.Y - equipLine.EndPoint.Y) / 2;
                            colorMark.Location = new Point3d(x, y, 0);
                            colorMark.Contents = units[index[k]].colors[color];
                            colorMark.Rotation = 1.5708;
                            colorMark.Attachment = AttachmentPoint.BottomCenter;
                            modSpace.AppendEntity(colorMark);
                            acTrans.AddNewlyCreatedDBObject(colorMark, true);

                            Polyline equipTerm = new Polyline();
                            equipTerm.SetDatabaseDefaults();
                            equipTerm.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            equipTerm.Closed = true;
                            equipTerm.AddVertexAt(0, new Point2d(equipLine.EndPoint.X - 2.5, equipLine.EndPoint.Y), 0, 0, 0);
                            equipTerm.AddVertexAt(1, equipTerm.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                            equipTerm.AddVertexAt(2, equipTerm.GetPoint2dAt(1).Add(new Vector2d(0, -5)), 0, 0, 0);
                            equipTerm.AddVertexAt(3, equipTerm.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                            modSpace.AppendEntity(equipTerm);
                            acTrans.AddNewlyCreatedDBObject(equipTerm, true);

                            MText cableTextDown = new MText();
                            cableTextDown.SetDatabaseDefaults();
                            cableTextDown.TextHeight = 2.5;
                            if (tst.Has("spds 2.5-0.85"))
                                cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                            cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            x = equipTerm.GetPoint3dAt(0).X + (equipTerm.GetPoint3dAt(1).X - equipTerm.GetPoint3dAt(0).X) / 2;
                            y = equipTerm.GetPoint3dAt(2).Y + (equipTerm.GetPoint3dAt(1).Y - equipTerm.GetPoint3dAt(2).Y) / 2;
                            cableTextDown.Location = new Point3d(x, y, 0);
                            cableTextDown.Attachment = AttachmentPoint.MiddleCenter;
                            cableTextDown.Contents = units[index[k]].equipTerminals[c];
                            modSpace.AppendEntity(cableTextDown);
                            acTrans.AddNewlyCreatedDBObject(cableTextDown, true);

                            lowestPoint = equipTerm.GetPoint3dAt(3).Y;
                            if (color < units[index[k]].colors.Count - 1) color++;
                            else color = 0;
                        }

                        Polyline equipFrame = new Polyline();
                        equipFrame.SetDatabaseDefaults();
                        if (lineTypeTable.Has("штриховая2"))
                            equipFrame.LinetypeId = lineTypeTable["штриховая2"];
                        else if (lineTypeTable.Has("hidden2"))
                            equipFrame.LinetypeId = lineTypeTable["hidden2"];
                        equipFrame.LinetypeScale = 5;
                        equipFrame.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        equipFrame.Closed = true;
                        equipFrame.AddVertexAt(0, new Point2d(equipJumper.StartPoint.X - 10, equipJumper.EndPoint.Y + 5), 0, 0, 0);
                        equipFrame.AddVertexAt(1, new Point2d(equipJumper.EndPoint.X + 10, equipJumper.EndPoint.Y + 5), 0, 0, 0);
                        equipFrame.AddVertexAt(2, new Point2d(equipJumper.EndPoint.X + 10, lowestPoint - 4), 0, 0, 0);
                        equipFrame.AddVertexAt(3, new Point2d(equipJumper.StartPoint.X - 10, lowestPoint - 4), 0, 0, 0);
                        modSpace.AppendEntity(equipFrame);
                        acTrans.AddNewlyCreatedDBObject(equipFrame, true);

                        double C = equipFrame.GetPoint2dAt(0).X;
                        double D = equipFrame.GetPoint2dAt(1).X;
                        double A = tables[curTableNum].Position.X;
                        double B = tables[curTableNum].Position.X + tables[curTableNum].Width;
                        width = 2 * (C - B) + (D - C);
                        tables[curTableNum].InsertColumns(tables[curTableNum].Columns.Count, width, 1);
                        tables[curTableNum].UnmergeCells(tables[curTableNum].Rows[0]);
                        if (tst.Has("spds 2.5-0.85"))
                            tables[curTableNum].Columns[tables[curTableNum].Columns.Count - 1].TextStyleId = tst["spds 2.5-0.85"];
                        tables[curTableNum].Columns[tables[curTableNum].Columns.Count - 1].TextHeight = 2.5;
                        tables[curTableNum].Cells[0, tables[curTableNum].Columns.Count - 1].TextString = units[index[k]].equipType == "" ? "-" : units[index[k]].equipType;
                        tables[curTableNum].Cells[1, tables[curTableNum].Columns.Count - 1].TextString = units[index[k]].designation;
                        tables[curTableNum].Cells[2, tables[curTableNum].Columns.Count - 1].TextString = units[index[k]].param;
                        tables[curTableNum].Cells[3, tables[curTableNum].Columns.Count - 1].TextString = units[index[k]].equipment;
                        tables[curTableNum].GenerateLayout();
                    }

                    prevTbox = units[i].tBoxName;
                    points = new List<Point3d>();
                    index = new List<int>();

                    cableLineUp = new Line();
                    cableLineUp.SetDatabaseDefaults();
                    cableLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cableLineUp.StartPoint = units[i].cableOutput[0];
                    cableLineUp.EndPoint = cableLineUp.StartPoint;
                    modSpace.AppendEntity(cableLineUp);
                    acTrans.AddNewlyCreatedDBObject(cableLineUp, true);
                    for (int j = 0; j < units[i].cableOutput.Count; j++)
                    {
                        cableLineUp.EndPoint = units[i].cableOutput[j];
                        points.Add(units[i].cableOutput[j]);
                        index.Add(i);
                    }
                }
                if (units[i].newSheet) curTableNum++;
            }
        }

        private static Point3d drawTerminalBoxUnit(Transaction acTrans, BlockTableRecord modSpace, int index, Point3d point, Point3d cableEndPoint, out Point3d lastTerminalPoint)
        {
            Line tBoxInput = new Line();
            tBoxInput.SetDatabaseDefaults();
            tBoxInput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInput.StartPoint = new Point3d(point.X, cableEndPoint.Y, 0);
            tBoxInput.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(tBoxInput);
            acTrans.AddNewlyCreatedDBObject(tBoxInput, true);

            Line tBoxInputBranch = new Line();
            tBoxInputBranch.SetDatabaseDefaults();
            tBoxInputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInputBranch.StartPoint = tBoxInput.StartPoint.Add(new Vector3d(0, -8, 0));
            tBoxInputBranch.EndPoint = tBoxInputBranch.StartPoint.Add(new Vector3d(0, 0, 0));
            modSpace.AppendEntity(tBoxInputBranch);
            acTrans.AddNewlyCreatedDBObject(tBoxInputBranch, true);

            double lowestPoint = 0;
            Point3d lastTermPoint = new Point3d();
            int color = 0;
            for (int i = units[index].colors.Count - 1; i >= 0; i--)
            {
                tBoxInputBranch.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(-5 * i, -8, 0));

                Line tBoxBranchDown = new Line();
                tBoxBranchDown.SetDatabaseDefaults();
                tBoxBranchDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxBranchDown.StartPoint = tBoxInputBranch.EndPoint;
                tBoxBranchDown.EndPoint = tBoxBranchDown.StartPoint.Add(new Vector3d(0, -8, 0));
                modSpace.AppendEntity(tBoxBranchDown);
                acTrans.AddNewlyCreatedDBObject(tBoxBranchDown, true);

                Polyline cablePolyUp = new Polyline();
                cablePolyUp.SetDatabaseDefaults();
                cablePolyUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cablePolyUp.Closed = true;
                cablePolyUp.AddVertexAt(0, new Point2d(tBoxBranchDown.EndPoint.X - 2.5, tBoxBranchDown.EndPoint.Y), 0, 0, 0);
                cablePolyUp.AddVertexAt(1, cablePolyUp.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                cablePolyUp.AddVertexAt(2, cablePolyUp.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
                cablePolyUp.AddVertexAt(3, cablePolyUp.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
                modSpace.AppendEntity(cablePolyUp);
                acTrans.AddNewlyCreatedDBObject(cablePolyUp, true);

                MText cableTextUp = new MText();
                cableTextUp.SetDatabaseDefaults();
                cableTextUp.TextHeight = 2.5;
                if (tst.Has("spds 2.5-0.85"))
                    cableTextUp.TextStyleId = tst["spds 2.5-0.85"];
                cableTextUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                double x = cablePolyUp.GetPoint3dAt(0).X + (cablePolyUp.GetPoint3dAt(1).X - cablePolyUp.GetPoint3dAt(0).X) / 2;
                double y = cablePolyUp.GetPoint3dAt(2).Y + (cablePolyUp.GetPoint3dAt(1).Y - cablePolyUp.GetPoint3dAt(2).Y) / 2;
                cableTextUp.Location = new Point3d(x, y, 0);
                cableTextUp.Contents = units[index].designation + "/" + curTermNumber;
                cableTextUp.Rotation = 1.5708;
                cableTextUp.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(cableTextUp);
                acTrans.AddNewlyCreatedDBObject(cableTextUp, true);

                Line colorLineUp = new Line();
                colorLineUp.SetDatabaseDefaults();
                colorLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                colorLineUp.StartPoint = cablePolyUp.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                colorLineUp.EndPoint = colorLineUp.StartPoint.Add(new Vector3d(0, -8, 0));
                modSpace.AppendEntity(colorLineUp);
                acTrans.AddNewlyCreatedDBObject(colorLineUp, true);

                MText colorMarkUp = new MText();
                colorMarkUp.SetDatabaseDefaults();
                colorMarkUp.TextHeight = 2.5;
                if (tst.Has("spds 2.5-0.85"))
                    colorMarkUp.TextStyleId = tst["spds 2.5-0.85"];
                colorMarkUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = colorLineUp.EndPoint.X - 1;
                y = colorLineUp.EndPoint.Y + (colorLineUp.StartPoint.Y - colorLineUp.EndPoint.Y) / 2;
                colorMarkUp.Location = new Point3d(x, y, 0);
                colorMarkUp.Contents = units[index].colors[color];
                colorMarkUp.Rotation = 1.5708;
                colorMarkUp.Attachment = AttachmentPoint.BottomCenter;
                modSpace.AppendEntity(colorMarkUp);
                acTrans.AddNewlyCreatedDBObject(colorMarkUp, true);

                Polyline termPoly = new Polyline();
                termPoly.SetDatabaseDefaults();
                termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                termPoly.Closed = true;
                termPoly.AddVertexAt(0, new Point2d(colorLineUp.EndPoint.X - 2.5, colorLineUp.EndPoint.Y), 0, 0, 0);
                termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(1).Add(new Vector2d(0, -10)), 0, 0, 0);
                termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(0).Add(new Vector2d(0, -10)), 0, 0, 0);
                modSpace.AppendEntity(termPoly);
                acTrans.AddNewlyCreatedDBObject(termPoly, true);
                lastTermPoint = termPoly.GetPoint3dAt(1);

                MText termText = new MText();
                termText.SetDatabaseDefaults();
                termText.TextHeight = 2.5;
                if (tst.Has("spds 2.5-0.85"))
                    termText.TextStyleId = tst["spds 2.5-0.85"];
                termText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2;
                y = termPoly.GetPoint3dAt(2).Y + (termPoly.GetPoint3dAt(1).Y - termPoly.GetPoint3dAt(2).Y) / 2;
                termText.Location = new Point3d(x, y, 0);
                termText.Contents = "";
                termText.Rotation = 1.5708;
                termText.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(termText);
                acTrans.AddNewlyCreatedDBObject(termText, true);

                Line colorLineDown = new Line();
                colorLineDown.SetDatabaseDefaults();
                colorLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                colorLineDown.StartPoint = termPoly.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                colorLineDown.EndPoint = colorLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                modSpace.AppendEntity(colorLineDown);
                acTrans.AddNewlyCreatedDBObject(colorLineDown, true);

                MText colorMarkDown = new MText();
                colorMarkDown.SetDatabaseDefaults();
                colorMarkDown.TextHeight = 2.5;
                if (tst.Has("spds 2.5-0.85"))
                    colorMarkDown.TextStyleId = tst["spds 2.5-0.85"];
                colorMarkDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = colorLineDown.EndPoint.X - 1;
                y = colorLineDown.EndPoint.Y + (colorLineDown.StartPoint.Y - colorLineDown.EndPoint.Y)/2;
                colorMarkDown.Location = new Point3d(x, y, 0);
                colorMarkDown.Contents = units[index].colors[color];
                colorMarkDown.Rotation = 1.5708;
                colorMarkDown.Attachment = AttachmentPoint.BottomCenter;
                modSpace.AppendEntity(colorMarkDown);
                acTrans.AddNewlyCreatedDBObject(colorMarkDown, true);

                Polyline cablePolyDown = new Polyline();
                cablePolyDown.SetDatabaseDefaults();
                cablePolyDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cablePolyDown.Closed = true;
                cablePolyDown.AddVertexAt(0, new Point2d(colorLineDown.EndPoint.X - 2.5, colorLineDown.EndPoint.Y), 0, 0, 0);
                cablePolyDown.AddVertexAt(1, cablePolyDown.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                cablePolyDown.AddVertexAt(2, cablePolyDown.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
                cablePolyDown.AddVertexAt(3, cablePolyDown.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
                modSpace.AppendEntity(cablePolyDown);
                acTrans.AddNewlyCreatedDBObject(cablePolyDown, true);

                MText cableTextDown = new MText();
                cableTextDown.SetDatabaseDefaults();
                cableTextDown.TextHeight = 2.5;
                if (tst.Has("spds 2.5-0.85"))
                    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = cablePolyDown.GetPoint3dAt(0).X + (cablePolyDown.GetPoint3dAt(1).X - cablePolyDown.GetPoint3dAt(0).X) / 2;
                y = cablePolyDown.GetPoint3dAt(2).Y + (cablePolyDown.GetPoint3dAt(1).Y - cablePolyDown.GetPoint3dAt(2).Y) / 2;
                cableTextDown.Location = new Point3d(x, y, 0);
                cableTextDown.Contents = units[index].designation + "/" + curTermNumber;
                cableTextDown.Rotation = 1.5708;
                cableTextDown.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(cableTextDown);
                acTrans.AddNewlyCreatedDBObject(cableTextDown, true);

                Line tBoxBranchOutput = new Line();
                tBoxBranchOutput.SetDatabaseDefaults();
                tBoxBranchOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxBranchOutput.StartPoint = cablePolyDown.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                tBoxBranchOutput.EndPoint = tBoxBranchOutput.StartPoint.Add(new Vector3d(0, -8, 0));
                modSpace.AppendEntity(tBoxBranchOutput);
                acTrans.AddNewlyCreatedDBObject(tBoxBranchOutput, true);

                lowestPoint = tBoxBranchOutput.EndPoint.Y;
                curTermNumber++;
                if (color < units[index].colors.Count - 1) color++;
                else color = 0;
            }
            

            tBoxInputBranch.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(-5 * (units[index].colors.Count - 1), -8, 0));
            Line tBoxOutputBranch = new Line();
            tBoxOutputBranch.SetDatabaseDefaults();
            tBoxOutputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxOutputBranch.StartPoint = new Point3d(tBoxInputBranch.StartPoint.X, lowestPoint, 0);
            tBoxOutputBranch.EndPoint = tBoxOutputBranch.StartPoint.Add(new Vector3d(-5 * (units[index].colors.Count - 1), 0, 0));
            modSpace.AppendEntity(tBoxOutputBranch);
            acTrans.AddNewlyCreatedDBObject(tBoxOutputBranch, true);

            lastTerminalPoint = drawTBoxGnd(acTrans, modSpace, lastTermPoint, tBoxInputBranch.EndPoint.X, tBoxInputBranch.StartPoint.X, tBoxInputBranch.StartPoint.Y, tBoxOutputBranch.StartPoint.Y);

            Line tBoxOutput = new Line();
            tBoxOutput.SetDatabaseDefaults();
            tBoxOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxOutput.StartPoint = tBoxOutputBranch.StartPoint;
            tBoxOutput.EndPoint = tBoxOutput.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(tBoxOutput);
            acTrans.AddNewlyCreatedDBObject(tBoxOutput, true);

            MText pairMark = new MText();
            curPairMark = pairMark;
            pairMark.SetDatabaseDefaults();
            pairMark.TextHeight = 2.5;
            if (tst.Has("spds 2.5-0.85"))
                pairMark.TextStyleId = tst["spds 2.5-0.85"];
            pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double X = tBoxOutput.EndPoint.X - 1;
            double Y = tBoxOutput.EndPoint.Y + (tBoxOutput.StartPoint.Y - tBoxOutput.EndPoint.Y) / 2;
            pairMark.Location = new Point3d(X, Y, 0);
            pairMark.Contents = "PR" + (curPairNumber < 10 ? "0" + curPairNumber.ToString() : curPairNumber.ToString());
            pairMark.Rotation = 1.5708;
            pairMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(pairMark);
            acTrans.AddNewlyCreatedDBObject(pairMark, true);

            return tBoxOutput.EndPoint;
        }

        private static void insertSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d point)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            string filename;
            string blockName;
            if (currentSheetNumber==1)
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
                        BlockTableRecord btrcrd = br.BlockTableRecord.GetObject(OpenMode.ForWrite) as BlockTableRecord;
                        foreach (ObjectId id in btrcrd)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForWrite);
                            AttributeDefinition attDef = obj as AttributeDefinition;
                            if ((attDef != null) && (!attDef.Constant))
                            {
                                #region attributes
                                switch (attDef.Tag)
                                {
                                    case "SH":
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                attRef.TextString = currentSheetNumber.ToString();
                                                currentSheetNumber++;
                                                br.AttributeCollection.AppendAttribute(attRef);
                                                acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            }
                                            break;
                                        }
                                }
                                #endregion
                            }
                        }
                    }
                }
                else editor.WriteMessage("В файле не найден блок с именем \"[{0}\"", blockName);
            }
        }

        private static void loadFonts(Transaction acTrans, BlockTableRecord modSpace, Database acdb)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            ObjectIdCollection id = new ObjectIdCollection();
            string filename = @"Data\Font.dwg";
            using (Database sourceDb = new Database(false, true))
            {
                if (System.IO.File.Exists(filename))
                {
                    sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                    using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                    {
                        TextStyleTable table = (TextStyleTable)trans.GetObject(sourceDb.TextStyleTableId, OpenMode.ForRead);
                        if (table.Has("spds 2.5-0.85"))
                            id.Add(table["spds 2.5-0.85"]);
                        trans.Commit();
                    }
                }
                else editor.WriteMessage("Не найден файл {0}", filename);
                if (id.Count > 0)
                {
                    acTrans.TransactionManager.QueueForGraphicsFlush();
                    IdMapping iMap = new IdMapping();
                    acdb.WblockCloneObjects(id, acdb.TextStyleTableId, iMap, DuplicateRecordCloning.Replace, false);
                }
            }
        }
    }
}
