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
        private static List<tBoxUnit> tBoxes;
        private static Editor editor;
        private static Table currentTable = null;

        private static LinetypeTable lineTypeTable;
        private static TextStyleTable tst;
        private static Polyline currentTBox = null;
        private static Unit currentUnit;
        private static MText currentTBoxName = null;
        private static Leader currentLeader = null;
        private static double width = 0;
        private static List<Line> cableLineUps;
        private static List<MText> cableLineText;
        private static List<Table> tables = new List<Table>();
        private static int currentSheetNumber = 1;

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
            public List<string> colors;
            public List<string> equipTerminals;
            public Line cableMarkLine;
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
                cableMarkLine = null;
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
                currentSheetNumber = 1;
                units = loadData(file.FileName);
                tBoxes = new List<tBoxUnit>();
                tables = new List<Table>();
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
                        drawUnits(acTrans, acDb, acModSpace, units, drawSheet(acTrans, acModSpace, acDb, startPoint));
                        currentSheetNumber = 1;
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

            //Table table = new Table();
            //table.Position = shieldPoly.GetPoint3dAt(0).Add(new Vector3d(-30, -272, 0));
            //table.SetSize(4, 1);
            //table.TableStyle = acdb.Tablestyle;

            //table.SetTextHeight(0, 0, 2.5);
            //table.Cells[0, 0].TextString = "Тип оборудования";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[0, 0].TextStyleId = tst["spds 2.5-0.85"];
            //table.SetAlignment(0, 0, CellAlignment.MiddleCenter);
            //table.Rows[0].IsMergeAllEnabled = false;

            //table.SetTextHeight(1, 0, 2.5);
            //table.Cells[1, 0].TextString = "Обозначение по проекту";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[1, 0].TextStyleId = tst["spds 2.5-0.85"];
            //table.SetAlignment(1, 0, CellAlignment.MiddleCenter);

            //table.SetTextHeight(2, 0, 2.5);
            //table.Cells[2, 0].TextString = "Параметры";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[2, 0].TextStyleId = tst["spds 2.5-0.85"];
            //table.SetAlignment(2, 0, CellAlignment.MiddleCenter);

            //table.SetTextHeight(3, 0, 2.5);
            //table.Cells[3, 0].TextString = "Оборудоание";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[3, 0].TextStyleId = tst["spds 2.5-0.85"];
            //table.SetAlignment(3, 0, CellAlignment.MiddleCenter);

            //table.Columns[0].Width = 30;
            //table.GenerateLayout();
            //currentTable = table;
            //tables.Add(table);

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
                    currentTBoxName = null;
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
                    int ind = 0;
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

                            prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal, 15, out lowestPoint, units[j].designation, i + 1);
                            leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                            ind++;
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
                                currentTBoxName = null;
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
            cableLineUp.EndPoint = cableLineUp.StartPoint.Add(new Vector3d(0, -44, 0));
            modSpace.AppendEntity(cableLineUp);
            acTrans.AddNewlyCreatedDBObject(cableLineUp, true);
           
            MText textLine = new MText();
            textLine.SetDatabaseDefaults();
            textLine.TextHeight = 3;
            if (tst.Has("spds 2.5-0.85"))
                textLine.TextStyleId = tst["spds 2.5-0.85"];
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
            double Y = cableName.GetPoint2dAt(2).Y + (cableName.GetPoint2dAt(1).Y - cableName.GetPoint2dAt(2).Y) / 2;
            Point3d cableNameCentre = new Point3d(X, Y, 0);
            MText textName = new MText();
            textName.SetDatabaseDefaults();
            textName.TextHeight = 2.5;
            if (tst.Has("spds 2.5-0.85"))
                textName.TextStyleId = tst["spds 2.5-0.85"];
            textName.Location = cableNameCentre;
            textName.Attachment = AttachmentPoint.MiddleCenter;
            textName.Width = 25;
            textName.Contents = unit.tBoxName == string.Empty ? unit.cupboardName.Split(' ')[1] + "/" + unit.designation : unit.cupboardName.Split(' ')[1] + "/" + unit.tBoxName;
            modSpace.AppendEntity(textName);
            acTrans.AddNewlyCreatedDBObject(textName, true);
            if (cableName.GetPoint2dAt(1).X - cableName.GetPoint2dAt(0).X < textName.ActualWidth)
            {
                Y = cableName.GetPoint2dAt(0).Y - 2;
                cableNameCentre = new Point3d(X, Y, 0);
                textName.Location = cableNameCentre;
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
            cableMark.TextHeight = 3;
            if (tst.Has("spds 2.5-0.85"))
                cableMark.TextStyleId = tst["spds 2.5-0.85"];
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

        private static void drawOther(Transaction acTrans, BlockTableRecord modSpace, int ind)
        {
            Line cableLine2 = units[ind].cableMarkLine;

            if (units[ind].tBoxName != string.Empty)
            {
                Unit unit = units[ind];
                unit.cableMarkLine = drawTerminalBox(acTrans, modSpace, cableLine2, ind);
                units[ind] = unit;
            }
            else
            {
                currentTBoxName = null;
            }
        }

        private static void drawEquip(Transaction acTrans, BlockTableRecord modSpace, int ind)
        {
            double leftEdgeX = 0;
            double rightEdgeX = 0;
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = units[ind].cableMarkLine.EndPoint.X - ((units[ind].equipTerminals.Count - 1) * 5) / 2;
            jumperLineDown.StartPoint = new Point3d(x, units[ind].cableMarkLine.EndPoint.Y, 0);
            x = x + ((units[ind].equipTerminals.Count - 1) * 5);
            jumperLineDown.EndPoint = new Point3d(x, units[ind].cableMarkLine.EndPoint.Y, 0);
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
                textLine.TextHeight = 3;
                if (tst.Has("spds 2.5-0.85"))
                    textLine.TextStyleId = tst["spds 2.5-0.85"];
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
                text.Height = 3;
                if (tst.Has("spds 2.5-0.85"))
                    text.TextStyleId = tst["spds 2.5-0.85"];
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
            if (tst.Has("spds 2.5-0.85"))
                currentTable.Columns[currentTable.Columns.Count - 1].TextStyleId = tst["spds 2.5-0.85"];
            currentTable.Columns[currentTable.Columns.Count - 1].TextHeight = 2.5;
            currentTable.Cells[0, currentTable.Columns.Count - 1].TextString = units[ind].equipType == "" ? "-" : units[ind].equipType;
            currentTable.Cells[1, currentTable.Columns.Count - 1].TextString = units[ind].designation;
            currentTable.Cells[2, currentTable.Columns.Count - 1].TextString = units[ind].param;
            currentTable.Cells[3, currentTable.Columns.Count - 1].TextString = units[ind].equipment;
            currentTable.GenerateLayout();
        }

        private static Line drawTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Line cableLine, int ind)
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
                textLine.TextHeight = 3;
                if (tst.Has("spds 2.5-0.85"))
                    textLine.TextStyleId = tst["spds 2.5-0.85"];
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
                text.Height = 3;
                if (tst.Has("spds 2.5-0.85"))
                    text.TextStyleId = tst["spds 2.5-0.85"];
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
                textLine.TextHeight = 3;
                if (tst.Has("spds 2.5-0.85"))
                    textLine.TextStyleId = tst["spds 2.5-0.85"];
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
            units[ind] = unit;

            if (currentTBoxName == null || currentUnit.tBoxName != units[ind].tBoxName || units[ind].newSheet)
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
                tBoxText.Height = 3;
                if (tst.Has("spds 2.5-0.85"))
                    tBoxText.TextStyleId = tst["spds 2.5-0.85"];
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
                tBoxName.TextHeight = 3;
                if (tst.Has("spds 2.5-0.85"))
                    tBoxName.TextStyleId = tst["spds 2.5-0.85"];
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
                tBoxName.TextHeight = 3;
                if (tst.Has("spds 2.5-0.85"))
                    tBoxName.TextStyleId = tst["spds 2.5-0.85"];
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

            return cLineOutput;
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
