using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace AcElectricalSchemePlugin
{
    static class ConnectionSchemeClass
    {
        private static Document acDoc;
        private static List<Unit> units;
        private static Editor editor;
        private static int curTableNum;
        private static LinetypeTable lineTypeTable;
        //private static TextStyleTable tst;
        private static double width = 0;
        private static List<Table> tables = new List<Table>();
        private static int currentSheetNumber = 1;
        private static List<tBox> tboxes;
        private static int curTermNumber;
        private static int curPairNumber;
        private static MText curPairMark;
        private static double rigthEdgeXTable;
        private static List<Group> groups;
        private static Polyline curDefGnd;
        private static double widthFactor = 0.5;

        struct tBox
        {
            public string Name;
            public int Count;
            public int LastTerminalNumber;
            public int LastShieldNumber;
            public int LastPairNumber;
            public List<MText> textName;
            public List<Leader> ldr;

            public bool shield;
            public Point3d framePoint;
            public Point3d lastTermPoint;
            public double cableLineSPLX;
            public double cableLineSPRX;
            public double cableLineEPY;

            public double gndL;
            public double gndR;
            public double gndLowP;
            public double gndY;
            public Polyline gndShield;
        }

        struct Unit
        {
            public string cupboardName;
            public string tBoxName;
            public string designation;
            public int numOfGear;
            public string linkText;
            public string param;
            public string equipment;
            public string equipType;
            public string cableMark;
            public bool shield;
            public List<string> terminals;
            public List<string> colors;
            public List<string> equipTerminals;
            public List<Line> cableOutput;
            public Point3d equipPoint;
            public List<Line> cables;
            public MText cableName;
            public bool newSheet;
            public Unit(string _cupboardName, string _tBoxName, string _designation, int  _numOfGear, string _linkText, string _param, string _equipment, string _equipType, string _cableMark, bool _shield, List<string> _terminals, List<string> _colors, List<string> _equipTerminals)
            {
                cupboardName = _cupboardName;
                tBoxName = _tBoxName;
                designation = _designation;
                numOfGear = _numOfGear;
                linkText = _linkText;
                param = _param;
                equipment = _equipment;
                equipType = _equipType;
                cableMark = _cableMark;
                shield = _shield;
                terminals = _terminals;
                colors = _colors;
                equipTerminals = _equipTerminals;
                cableOutput = new List<Line>();
                cables = new List<Line>();
                cableName = null;
                newSheet = false;
                equipPoint = new Point3d();
            }
        }

        struct Group
        {
            public List<Unit> Units;
            public string GroupName;
            public int CableNum;

            public Group(List<Unit> _units, string groupName, int cableNum)
            {
                Units = _units;
                GroupName = groupName;
                CableNum = cableNum;
            }

            public void addUnit(Unit unit)
            {
                Units.Add(unit);
                CableNum += 1;
            }
        }

        private static List<Unit> LoadData(string path)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;MAXSCANROWS=0'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetSchema("Tables", new string[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [Full$]");
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.FillSchema(dataSet, SchemaType.Source);
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
                int.TryParse(dataSet.Tables[0].Rows[row][29].ToString(), out unit.numOfGear);
                unit.linkText = dataSet.Tables[0].Rows[row][30].ToString();
                unit.equipTerminals = equipTerminals;
                units.Add(unit);
            }
            return units;
        }

        private static List<Group> FindGroups(List<Unit> Units)
        {
            List<Group> Groups = new List<Group>();

            Group group = new Group(new List<Unit>(), Units[0].tBoxName, 0);
            group.addUnit(Units[0]);
            Groups.Add(group);

            for (int i = 1; i < Units.Count; i++)
            {
                if (Units[i].tBoxName == Groups[Groups.Count - 1].Units[0].tBoxName && Units[i].tBoxName != String.Empty && !Units[i].newSheet)
                {
                    group = Groups[Groups.Count - 1];
                    group.addUnit(Units[i]);
                    Groups[Groups.Count - 1] = group;
                }
                else
                {
                    group = new Group(new List<Unit>(), Units[i].tBoxName, 0);
                    group.addUnit(Units[i]);
                    Groups.Add(group);
                }
            }
            return Groups;
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
                rigthEdgeXTable = 0;
                width = 0;
                tables = new List<Table>();
                currentSheetNumber = 1;
                tboxes = new List<tBox>();
                curTermNumber = 1;
                curPairNumber = 1;
                currentSheetNumber = 1;
                units = LoadData(file.FileName);
                ConnectionScheme();
            }
        }

        private static void ConnectionScheme()
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
                        //tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForRead);
                        //TextStyleTableRecord tstr = null;
                        //if (tst.Has("spds 2.5-0.85"))
                        //{
                        //    tstr = (TextStyleTableRecord)acTrans.GetObject(tst["spds 2.5-0.85"], OpenMode.ForRead);
                        //    widthFactor = tstr.XScale;
                        //}
                        //loadFonts(acTrans, acModSpace, acDb);
                       
                        Point3d startPoint = selectedPoint.Value;
                        Polyline sh = DrawSheet(acTrans, acModSpace, acDb, startPoint);
                        DrawUnits(acTrans, acDb, acModSpace, units, sh);
                        DrawCables(acTrans, acModSpace, units, sh);
                        for (int i = 0; i < tables.Count; i++)
                        {
                            if (!acBlkTbl.Has(tables[i].Id))
                            {
                                acModSpace.AppendEntity(tables[i]);
                                acTrans.AddNewlyCreatedDBObject(tables[i], true);
                            }
                        }
                        RenameTBoxes(acTrans, acModSpace);
                        acDoc.Editor.Regen();
                        acTrans.Commit();
                    }
                }
            }
        }

        private static void RenameTBoxes(Transaction acTrans, BlockTableRecord modSpace)
        {
            for (int i = 0; i < tboxes.Count; i++)
            {
                if (tboxes[i].Name != String.Empty)
                {
                    if (tboxes[i].Count > 1)
                    {
                        for (int j = 0; j < tboxes[i].Count; j++)
                        {
                            MText text = tboxes[i].textName[j];
                            text.Contents += "." + (j + 1);
                            if (text.Contents.Contains("*"))
                            {
                                text.Contents = text.Contents.Replace("*", string.Empty);
                                text.Contents += "*";
                            }
                            tboxes[i].textName[j] = text;

                            Leader ldr = tboxes[i].ldr[j];
                            Point3d point = ldr.VertexAt(1);
                            ldr.SetVertexAt(2, new Point3d(point.X + text.ActualWidth + 1, point.Y, 0));
                        }
                    }
                }
                if (tboxes[i].shield)
                {
                    DrawTBoxGnd(acTrans, modSpace, tboxes[i], tboxes[i].lastTermPoint, tboxes[i].cableLineSPLX, tboxes[i].cableLineSPRX, tboxes[i].cableLineEPY, tboxes[i].cableLineEPY, true, tboxes[i].framePoint);
                    if (tboxes[i].gndShield!=null)
                        DrawGnd(acTrans, modSpace, tboxes[i].gndL, tboxes[i].gndR, tboxes[i].gndY, tboxes[i].gndLowP, tboxes[i].gndShield);
                }
            }
        }

        private static Polyline DrawSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d prevSheet)
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

            InsertSheet(acTrans, modSpace, acdb, new Point3d(shieldPoly.GetPoint2dAt(0).X - 92, shieldPoly.GetPoint2dAt(0).Y + 10, 0));

            Polyline defGndPoly = new Polyline();
            defGndPoly.SetDatabaseDefaults();
            defGndPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            defGndPoly.Closed = true;
            defGndPoly.AddVertexAt(0, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -13)), 0, 0, 0);
            defGndPoly.AddVertexAt(1, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(754, -13)), 0, 0, 0);
            defGndPoly.AddVertexAt(2, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(754, -15)), 0, 0, 0);
            defGndPoly.AddVertexAt(3, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -15)), 0, 0, 0);
            modSpace.AppendEntity(defGndPoly);
            acTrans.AddNewlyCreatedDBObject(defGndPoly, true);
            curDefGnd = defGndPoly;

            DBText defGndText = new DBText();
            defGndText.SetDatabaseDefaults();
            //if (tst.Has("spds 2.5-0.85"))
            //{
            //    defGndText.TextStyleId = tst["spds 2.5-0.85"];
            //    defGndText.WidthFactor = widthFactor;
            //}
            defGndText.WidthFactor = widthFactor;
            defGndText.Height = 3;
            defGndText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            defGndText.Position = defGndPoly.GetPoint3dAt(0).Add(new Vector3d(0, 2, 0));
            defGndText.TextString = "Шина защитного заземления";
            defGndText.HorizontalMode = TextHorizontalMode.TextLeft;
            modSpace.AppendEntity(defGndText);
            acTrans.AddNewlyCreatedDBObject(defGndText, true);

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
            //if (tst.Has("spds 2.5-0.85"))
            //{
            //    text.TextStyleId = tst["spds 2.5-0.85"];
            //    text.WidthFactor = widthFactor;
            //}
            text.WidthFactor = widthFactor;
            text.Height = 3;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = gndPoly.GetPoint3dAt(0).Add(new Vector3d(0, 2, 0));
            text.TextString = "Шина функционального заземления";
            text.HorizontalMode = TextHorizontalMode.TextLeft;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            Table table = new Table();
            table.Position = shieldPoly.GetPoint3dAt(0).Add(new Vector3d(-30, -335, 0));
            table.SetSize(4, 1);
            table.TableStyle = acdb.Tablestyle;

            table.Cells[0,0].TextHeight = 2.5;
            table.Cells[0, 0].TextString = "Тип оборудования";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[0, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
            table.Rows[0].IsMergeAllEnabled = false;

            table.Cells[1, 0].TextHeight = 2.5;
            table.Cells[1, 0].TextString = "Обозначение по проекту";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[1, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;

            table.Cells[2, 0].TextHeight = 2.5;
            table.Cells[2, 0].TextString = "Параметры";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[2, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.Cells[2, 0].Alignment = CellAlignment.MiddleCenter;

            table.Cells[3, 0].TextHeight = 2.5;
            table.Cells[3, 0].TextString = "Оборудование";
            //if (tst.Has("spds 2.5-0.85"))
            //    table.Cells[3, 0].TextStyleId = tst["spds 2.5-0.85"];
            table.Cells[3, 0].Alignment = CellAlignment.MiddleCenter;

            table.Columns[0].Width = 30;
            table.GenerateLayout();
            tables.Add(table);

            return shieldPoly;
        }

        private static void DrawUnits(Transaction acTrans, Database acdb, BlockTableRecord modSpace, List<Unit> units, Polyline shield)
        {
            Polyline tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[0].cupboardName);
            Polyline prevPoly = tBoxPoly;
            Polyline prevTermPoly = null;
            string prevTerminal=string.Empty;
            string prevCupboard = units[0].cupboardName;
            bool newJ = false;
            for (int j = 0; j < units.Count; j++)
            {
                int color = 0;
                if (units[j].cupboardName.Contains("VA"))
                {
                    if (units[j].param.Contains("TS") && units[j].equipTerminals.Count == 4)
                    {
                        #region VAwithTS
                        List<Point3d> points = new List<Point3d>();
                        Unit un = units[j];
                        un.cableOutput = new List<Line>();
                        units[j] = un;
                        bool aborted = false;
                        Point3d lowestPoint = Point3d.Origin;
                        double l = 0;
                        double r = 0;
                        bool shielded = false;
                        if (prevCupboard != units[j].cupboardName)
                        {
                            shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                            tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                            for (int i = 0; i < units[j].terminals.Count; i += units[j].terminals.Count - 1)
                            {
                                string terminalTag = units[j].terminals[i].Split('[', ']')[1];
                                string terminal = units[j].terminals[i].Split('[', ']')[2];
                                if (prevTerminal == String.Empty)
                                {
                                    prevTerminal = terminalTag;
                                    DBText text = new DBText();
                                    text.SetDatabaseDefaults();
                                    text.Height = 3;
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
                                    text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                                    text.TextString = terminalTag;
                                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                                    text.AlignmentPoint = text.Position;
                                    modSpace.AppendEntity(text);
                                    acTrans.AddNewlyCreatedDBObject(text, true);
                                    newJ = false;
                                    shielded = false;
                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }
                                    
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
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
                                        aborted = true;
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : 6, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : 6, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }
                                    
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
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
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }
                                    
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                }
                            }
                            if (!shielded && !aborted)
                            {
                                r = prevTermPoly.GetPoint2dAt(1).X;

                                Line cableLine = new Line();
                                cableLine.SetDatabaseDefaults();
                                cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLine.StartPoint = points[0];
                                cableLine.EndPoint = points[points.Count - 1];
                                modSpace.AppendEntity(cableLine);
                                acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                Line cableLineDown = new Line();
                                cableLineDown.SetDatabaseDefaults();
                                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                modSpace.AppendEntity(cableLineDown);
                                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                units[j].cableOutput.Add(cableLineDown);
                                if (units[j].shield)
                                {
                                    DrawGnd(acTrans, modSpace, l - 3, r + 3, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);
                                }

                                points = new List<Point3d>();
                            }
                            if (!aborted) trans.Commit();
                            else j--;
                        }
                        newJ = true;
                        #endregion
                    }
                    else
                    {
                        #region VAwithoutTS
                        List<Point3d> points = new List<Point3d>();
                        Unit un = units[j];
                        un.cableOutput = new List<Line>();
                        units[j] = un;
                        bool aborted = false;
                        Point3d lowestPoint = Point3d.Origin;
                        double l = 0;
                        double r = 0;
                        bool shielded = false;
                        if (prevCupboard != units[j].cupboardName)
                        {
                            shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                            tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
                                    text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                                    text.TextString = terminalTag;
                                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                                    text.AlignmentPoint = text.Position;
                                    modSpace.AppendEntity(text);
                                    acTrans.AddNewlyCreatedDBObject(text, true);
                                    newJ = false;
                                    shielded = false;
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
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
                                        aborted = true;
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : 6, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
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
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(15, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    points.Add(lowestPoint);
                                    if (color < units[j].colors.Count - 1) color++;
                                    else color = 0;
                                    if (i == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                }
                            }
                            if (!shielded && !aborted)
                            {
                                r = prevTermPoly.GetPoint2dAt(1).X;

                                Line cableLine = new Line();
                                cableLine.SetDatabaseDefaults();
                                cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLine.StartPoint = points[0];
                                cableLine.EndPoint = points[points.Count - 1];
                                modSpace.AppendEntity(cableLine);
                                acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                Line cableLineDown = new Line();
                                cableLineDown.SetDatabaseDefaults();
                                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                modSpace.AppendEntity(cableLineDown);
                                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                units[j].cableOutput.Add(cableLineDown);
                                if (units[j].shield)
                                {
                                    DrawGnd(acTrans, modSpace, l - 3, r + 3, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);
                                }

                                points = new List<Point3d>();
                            }
                            if (!aborted) trans.Commit();
                            else j--;
                        }
                        newJ = true;
                        #endregion
                    }
                }
                else
                {
                    if (units[j].param.Contains("TS") && units[j].equipTerminals.Count == 4)
                    {
                        #region nonVAwithTS
                        List<Point3d> points = new List<Point3d>();
                        Unit un = units[j];
                        un.cableOutput = new List<Line>();
                        units[j] = un;
                        int offset = 15;
                        color = 0;
                        bool aborted = false;
                        double leftEdgeX = 0;
                        Point3d lowestPoint = Point3d.Origin;
                        double l = 0;
                        double r = 0;
                        bool shielded = false;
                        if (prevCupboard != units[j].cupboardName)
                        {
                            shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                            tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                            for (int i = 0; i < units[j].terminals.Count; i += units[j].terminals.Count - 1)
                            {
                                string terminalTag = units[j].terminals[i].Split('[', ']')[1];
                                string terminal = units[j].terminals[i].Split('[', ']')[2];
                                if (prevTerminal == String.Empty)
                                {
                                    prevTerminal = terminalTag;
                                    DBText text = new DBText();
                                    text.SetDatabaseDefaults();
                                    text.Height = 3;
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
                                    text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                                    text.TextString = terminalTag;
                                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                                    text.AlignmentPoint = text.Position;
                                    modSpace.AppendEntity(text);
                                    acTrans.AddNewlyCreatedDBObject(text, true);
                                    newJ = false;
                                    shielded = false;

                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }
                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

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
                                        p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(60, 0));
                                        prevPoly.SetPointAt(1, p1);
                                        p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(60, 0));
                                        prevPoly.SetPointAt(2, p2);
                                    }
                                    else
                                    {
                                        p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(offset, 0));
                                        prevPoly.SetPointAt(1, p1);
                                        p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(offset, 0));
                                        prevPoly.SetPointAt(2, p2);
                                    }
                                    if (p1.X >= shield.GetPoint2dAt(1).X - 170)
                                    {
                                        trans.Abort();
                                        aborted = true;
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : offset, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }

                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
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
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    if (i < units[j].terminals.Count - 1)
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    }
                                    else
                                    {
                                        prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, units[j].equipTerminals.Count);
                                    }
                                   
                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

                                        shielded = true;
                                    }
                                    if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                                }
                            }
                            if (!shielded && !aborted)
                            {
                                r = prevTermPoly.GetPoint2dAt(1).X;

                                Line cableLine = new Line();
                                cableLine.SetDatabaseDefaults();
                                cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLine.StartPoint = points[0];
                                cableLine.EndPoint = points[points.Count - 1];
                                modSpace.AppendEntity(cableLine);
                                acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                Line cableLineDown = new Line();
                                cableLineDown.SetDatabaseDefaults();
                                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                modSpace.AppendEntity(cableLineDown);
                                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                units[j].cableOutput.Add(cableLineDown);

                                DrawGnd(acTrans, modSpace, l - 3, r + 3, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                points = new List<Point3d>();
                            }
                            if (!aborted) trans.Commit();
                            else j--;
                        }
                        newJ = true;
                        #endregion
                    }
                    else
                    {
                        #region nonVAwithoutTS
                        List<Point3d> points = new List<Point3d>();
                        Unit un = units[j];
                        un.cableOutput = new List<Line>();
                        units[j] = un;
                        int offset = 15;
                        color = 0;
                        bool aborted = false;
                        double leftEdgeX = 0;
                        Point3d lowestPoint = Point3d.Origin;
                        double l = 0;
                        double r = 0;
                        bool shielded = false;
                        if (prevCupboard != units[j].cupboardName)
                        {
                            shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                            tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
                                    text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    text.Position = prevPoly.GetPoint3dAt(0).Add(new Vector3d(8, -6, 0));
                                    text.TextString = terminalTag;
                                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                                    text.AlignmentPoint = text.Position;
                                    modSpace.AppendEntity(text);
                                    acTrans.AddNewlyCreatedDBObject(text, true);
                                    newJ = false;
                                    shielded = false;
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    
                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

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
                                        p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(60, 0));
                                        prevPoly.SetPointAt(1, p1);
                                        p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(60, 0));
                                        prevPoly.SetPointAt(2, p2);
                                    }
                                    else
                                    {
                                        p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(offset, 0));
                                        prevPoly.SetPointAt(1, p1);
                                        p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(offset, 0));
                                        prevPoly.SetPointAt(2, p2);
                                    }
                                    if (p1.X >= shield.GetPoint2dAt(1).X - 170)
                                    {
                                        trans.Abort();
                                        aborted = true;
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevTermPoly.GetPoint2dAt(0).Add(new Vector2d(newJ ? 60 : offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    
                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

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
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //{
                                    //    text.TextStyleId = tst["spds 2.5-0.85"];
                                    //    text.WidthFactor = widthFactor;
                                    //}
                                    text.WidthFactor = widthFactor;
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
                                        shield = DrawSheet(acTrans, modSpace, acdb, shield.GetPoint3dAt(0).Add(new Vector3d(950, 0, 0)));
                                        tBoxPoly = DrawShieldTerminalBox(acTrans, modSpace, shield, units[j].cupboardName);
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
                                    prevTermPoly = DrawTerminal(acTrans, modSpace, prevPoly.GetPoint2dAt(0).Add(new Vector2d(offset, 0)), terminal, out lowestPoint, units[j], color, i + 1);
                                    
                                    points.Add(lowestPoint);
                                    if (color == 0) l = prevTermPoly.GetPoint2dAt(0).X;
                                    if (color < units[j].colors.Count - 1) { color++; offset = 6; }
                                    else
                                    {
                                        color = 0;
                                        offset = 15;
                                        r = prevTermPoly.GetPoint2dAt(1).X;

                                        Line cableLine = new Line();
                                        cableLine.SetDatabaseDefaults();
                                        cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLine.StartPoint = points[0];
                                        cableLine.EndPoint = points[points.Count - 1];
                                        modSpace.AppendEntity(cableLine);
                                        acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                        Line cableLineDown = new Line();
                                        cableLineDown.SetDatabaseDefaults();
                                        cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                        cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                        cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                        modSpace.AppendEntity(cableLineDown);
                                        acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                        units[j].cableOutput.Add(cableLineDown);

                                        if (units[j].tBoxName != null)
                                            if (units[j].tBoxName.Length > 0 && units[j].cableMark.Contains("ИЭ"))
                                                DrawGnd(acTrans, modSpace, l, r, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                        points = new List<Point3d>();

                                        shielded = true;
                                    }
                                    if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X;
                                }
                            }
                            if (!shielded && !aborted)
                            {
                                r = prevTermPoly.GetPoint2dAt(1).X;

                                Line cableLine = new Line();
                                cableLine.SetDatabaseDefaults();
                                cableLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLine.StartPoint = points[0];
                                cableLine.EndPoint = points[points.Count - 1];
                                modSpace.AppendEntity(cableLine);
                                acTrans.AddNewlyCreatedDBObject(cableLine, true);

                                Line cableLineDown = new Line();
                                cableLineDown.SetDatabaseDefaults();
                                cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                cableLineDown.StartPoint = new Point3d(points[0].X + (points[points.Count - 1].X - points[0].X) / 2, lowestPoint.Y, 0);
                                cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -8, 0));
                                modSpace.AppendEntity(cableLineDown);
                                acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                                units[j].cableOutput.Add(cableLineDown);

                                DrawGnd(acTrans, modSpace, l - 3, r + 3, prevTermPoly.GetPoint2dAt(2).Y, lowestPoint, shield);

                                points = new List<Point3d>();
                            }
                            if (!aborted) trans.Commit();
                            else j--;
                        }
                        newJ = true;
                        #endregion
                    }
                }
            }
            groups = FindGroups(units);
        }

        private static void DrawGnd(Transaction acTrans, BlockTableRecord modSpace, double leftEdgeX, double rightEdgeX, double Y, Point3d lowestPoint, Polyline shield)
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
        }

        private static void DrawGnd(Transaction acTrans, BlockTableRecord modSpace, double leftEdgeX, double rightEdgeX, double Y, double rightestPoint, Polyline shield)
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

        private static Point3d DrawTBoxGnd(Transaction acTrans, BlockTableRecord modSpace, tBox tbox, Point3d lastTermPoint, double leftEdgeX, double rightEdgeX, double upY, double downY,  bool osh, Point3d framePoint)
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
            //if (tst.Has("spds 2.5-0.85"))
            //    termText.TextStyleId = tst["spds 2.5-0.85"];
            termText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2;
            double y = termPoly.GetPoint3dAt(2).Y + (termPoly.GetPoint3dAt(1).Y - termPoly.GetPoint3dAt(2).Y) / 2;
            termText.Location = new Point3d(x, y, 0);
            if (upY != downY || !osh) termText.Contents = "\\W"+widthFactor+";S" + (tbox.LastShieldNumber < 10 ? "0" + tbox.LastShieldNumber.ToString() : tbox.LastShieldNumber.ToString());
            else
            {
                termText.Contents = "OSH";

                Line line1 = new Line();
                line1.SetDatabaseDefaults();
                line1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line1.StartPoint = framePoint;
                line1.EndPoint = line1.StartPoint.Add(new Vector3d(5, 0, 0));
                modSpace.AppendEntity(line1);
                acTrans.AddNewlyCreatedDBObject(line1, true);

                Line line2 = new Line();
                line2.SetDatabaseDefaults();
                line2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line2.StartPoint = line1.EndPoint;
                line2.EndPoint = line2.StartPoint.Add(new Vector3d(0, -3, 0));
                modSpace.AppendEntity(line2);
                acTrans.AddNewlyCreatedDBObject(line2, true);

                Line line3 = new Line();
                line3.SetDatabaseDefaults();
                line3.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line3.StartPoint = line2.EndPoint.Add(new Vector3d(-3, 0, 0));
                line3.EndPoint = line3.StartPoint.Add(new Vector3d(6, 0, 0));
                modSpace.AppendEntity(line3);
                acTrans.AddNewlyCreatedDBObject(line3, true);

                Line line4 = new Line();
                line4.SetDatabaseDefaults();
                line4.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line4.StartPoint = line2.EndPoint.Add(new Vector3d(-2, -1, 0));
                line4.EndPoint = line4.StartPoint.Add(new Vector3d(4, 0, 0));
                modSpace.AppendEntity(line4);
                acTrans.AddNewlyCreatedDBObject(line4, true);

                Line line5 = new Line();
                line5.SetDatabaseDefaults();
                line5.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                line5.StartPoint = line2.EndPoint.Add(new Vector3d(-1, -2, 0));
                line5.EndPoint = line5.StartPoint.Add(new Vector3d(2, 0, 0));
                modSpace.AppendEntity(line5);
                acTrans.AddNewlyCreatedDBObject(line5, true);
            }
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

        private static Polyline DrawShieldTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Polyline shield, string cupbordName)
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
            //if (tst.Has("spds 2.5-0.85"))
            //{
            //    text.TextStyleId = tst["spds 2.5-0.85"];
            //    text.WidthFactor = widthFactor;WidthFactor = widthFactor;
            //}
            text.WidthFactor = widthFactor;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = shield.GetPoint3dAt(0).Add(new Vector3d(382, -10, 0));
            text.TextString = cupbordName;
            text.Height = 4;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            return boxPoly;
        }

        private static Polyline DrawTerminal(Transaction acTrans, BlockTableRecord modSpace, Point2d point, string terminal, out Point3d lowestPoint, Unit unit, int color, int cableNumber)
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
            //if (tst.Has("spds 2.5-0.85"))
            //{
            //    text.TextStyleId = tst["spds 2.5-0.85"];
            //    text.WidthFactor = widthFactor;
            //}
            text.WidthFactor = widthFactor;
            text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            text.Position = termPoly.GetPoint3dAt(0).Add(new Vector3d(3, -4, 0));
            if (unit.param.ToLower() != "резерв")
            {
                text.TextString = terminal;
                text.Height = 3;
            }
            else
            {
                text.TextString = "Резерв";
                text.Height = 2;
            }
            text.Rotation = 1.5708;
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = text.Position;
            modSpace.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            if (unit.param.ToLower() == "резерв")
            {
                Line groundLine2 = new Line();
                groundLine2.SetDatabaseDefaults();
                if (lineTypeTable.Has("штриховая2"))
                    groundLine2.LinetypeId = lineTypeTable["штриховая2"];
                else if (lineTypeTable.Has("hidden2"))
                    groundLine2.LinetypeId = lineTypeTable["hidden2"];
                groundLine2.LinetypeScale = 5;
                groundLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLine2.StartPoint = new Point3d(termPoly.GetPoint2dAt(0).X+(termPoly.GetPoint2dAt(1).X-termPoly.GetPoint2dAt(0).X)/2, termPoly.GetPoint2dAt(0).Y, 0);
                groundLine2.EndPoint = new Point3d(groundLine2.StartPoint.X, curDefGnd.GetPoint2dAt(3).Y + 1, 0);
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

            Line cableLineUp = new Line();
            cableLineUp.SetDatabaseDefaults();
            cableLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp.StartPoint = new Point3d(termPoly.GetPoint2dAt(3).X + (termPoly.GetPoint2dAt(2).X - termPoly.GetPoint2dAt(3).X) / 2, termPoly.GetPoint2dAt(3).Y, 0);
            cableLineUp.EndPoint = cableLineUp.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(cableLineUp);
            acTrans.AddNewlyCreatedDBObject(cableLineUp, true);

            Polyline cablePoly = new Polyline();
            cablePoly.SetDatabaseDefaults();
            cablePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cablePoly.Closed = true;
            cablePoly.AddVertexAt(0, new Point2d(cableLineUp.EndPoint.X - 2.5, cableLineUp.EndPoint.Y), 0, 0, 0);
            cablePoly.AddVertexAt(1, cablePoly.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
            cablePoly.AddVertexAt(2, cablePoly.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
            cablePoly.AddVertexAt(3, cablePoly.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
            modSpace.AppendEntity(cablePoly);
            acTrans.AddNewlyCreatedDBObject(cablePoly, true);

            MText textLine = new MText();
            textLine.SetDatabaseDefaults();
            textLine.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    textLine.TextStyleId = tst["spds 2.5-0.85"];
            textLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = cablePoly.GetPoint3dAt(0).X + (cablePoly.GetPoint3dAt(1).X - cablePoly.GetPoint3dAt(0).X) / 2;
            double y = cablePoly.GetPoint3dAt(2).Y + (cablePoly.GetPoint3dAt(1).Y - cablePoly.GetPoint3dAt(2).Y) / 2;
            textLine.Location = new Point3d(x, y, 0);
            textLine.Contents = unit.param.ToLower() != "резерв" ? "\\W"+widthFactor+";"+unit.designation + "/" + cableNumber : "\\W"+widthFactor+";Резерв";
            textLine.Rotation = 1.5708;
            textLine.Attachment = AttachmentPoint.MiddleCenter;
            modSpace.AppendEntity(textLine);
            acTrans.AddNewlyCreatedDBObject(textLine, true);

            Line cableLineDown = new Line();
            cableLineDown.SetDatabaseDefaults();
            cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineDown.StartPoint = new Point3d(x, cablePoly.GetPoint2dAt(3).Y, 0);
            cableLineDown.EndPoint = cableLineDown.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(cableLineDown);
            acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

            MText colorMark = new MText();
            colorMark.SetDatabaseDefaults();
            colorMark.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    colorMark.TextStyleId = tst["spds 2.5-0.85"];
            colorMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            x = cableLineDown.EndPoint.X - 1;
            y = cableLineDown.EndPoint.Y + (cableLineDown.StartPoint.Y - cableLineDown.EndPoint.Y) / 2;
            colorMark.Location = new Point3d(x, y, 0);
            colorMark.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
            colorMark.Rotation = 1.5708;
            colorMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(colorMark);
            acTrans.AddNewlyCreatedDBObject(colorMark, true);

            lowestPoint = cableLineDown.EndPoint;
            return termPoly;
        }

        private static void DrawCables(Transaction acTrans, BlockTableRecord modSpace, List<Unit> units, Polyline shieldCupBoard)
        {
            for (int i = 0; i < groups.Count; i++)
            {
                tBox tbox;
                int index = 0;
                if (tboxes.Exists(x => x.Name == groups[i].GroupName) && groups[i].GroupName != String.Empty)
                {
                    index = tboxes.IndexOf(tboxes.Find(x => x.Name == groups[i].GroupName));
                    tbox = tboxes[index];
                    tbox.Count++;
                }
                else
                {
                    tbox = new tBox();
                    tbox.Name = groups[i].GroupName;
                    tbox.Count = 1;
                    tbox.LastPairNumber = 0;
                    tbox.LastShieldNumber = 1;
                    tbox.LastTerminalNumber = 1;
                    tbox.textName = new List<MText>();
                    tbox.ldr = new List<Leader>();
                    tboxes.Add(new tBox());
                    index = tboxes.Count - 1;
                }

                Line cableLineUp = new Line();
                cableLineUp.SetDatabaseDefaults();
                cableLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableLineUp.StartPoint = groups[i].Units[0].cableOutput[0].EndPoint;
                cableLineUp.EndPoint = groups[i].Units[groups[i].Units.Count - 1].cableOutput[groups[i].Units[groups[i].Units.Count - 1].cableOutput.Count - 1].EndPoint;
                modSpace.AppendEntity(cableLineUp);
                acTrans.AddNewlyCreatedDBObject(cableLineUp, true);

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
                //if (tst.Has("spds 2.5-0.85"))
                //    cableMark.TextStyleId = tst["spds 2.5-0.85"];
                cableMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                cableMark.Location = cableLine.EndPoint.Add(new Vector3d(-1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                cableMark.Contents = "\\W" + widthFactor + ";" + groups[i].Units[groups[i].Units.Count - 1].cableMark;
                cableMark.Rotation = 1.5708;
                cableMark.Attachment = AttachmentPoint.BottomCenter;
                modSpace.AppendEntity(cableMark);
                acTrans.AddNewlyCreatedDBObject(cableMark, true);

                MText textName = new MText();
                textName.SetDatabaseDefaults();
                textName.TextHeight = 3;
                //if (tst.Has("spds 2.5-0.85"))
                //    textName.TextStyleId = tst["spds 2.5-0.85"];
                textName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                textName.Location = cableLine.EndPoint.Add(new Vector3d(1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                textName.Attachment = AttachmentPoint.TopCenter;
                textName.Rotation = 1.5708;
                textName.Contents = groups[i].Units[groups[i].Units.Count - 1].tBoxName == string.Empty ? "\\W" + widthFactor + ";" + groups[i].Units[groups[i].Units.Count - 1].designation + "/" + groups[i].Units[groups[i].Units.Count - 1].cupboardName.Split(' ')[1] : "\\W" + widthFactor + ";" + groups[i].Units[groups[i].Units.Count - 1].tBoxName + "/" + groups[i].Units[groups[i].Units.Count - 1].cupboardName.Split(' ')[1];
                modSpace.AppendEntity(textName);
                acTrans.AddNewlyCreatedDBObject(textName, true);

                if (groups[i].Units[groups[i].Units.Count - 1].shield)
                {
                    tbox.gndL = cableLine.StartPoint.X - 6;
                    tbox.gndR = cableLine.StartPoint.X + 6;
                    tbox.gndY = cableLine.StartPoint.Y;
                    tbox.gndLowP = cableLineUp.EndPoint.X + 12.27;
                    tbox.gndShield = shieldCupBoard;
                }
                //if (groups[i].Units[groups[i].Units.Count - 1].shield && groups[i].Units.Count>1)
                //    drawGnd(acTrans, modSpace, cableLine.StartPoint.X - 6, cableLine.StartPoint.X + 6, cableLine.StartPoint.Y, cableLineUp.EndPoint.X + 12.27, shieldCupBoard);

                if (groups[i].Units[groups[i].Units.Count - 1].tBoxName != string.Empty)
                {
                    Line cableLineDown = new Line();
                    cableLineDown.SetDatabaseDefaults();
                    cableLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cableLineDown.StartPoint = new Point3d(cableLineUp.StartPoint.X, cableLine.EndPoint.Y, 0);
                    cableLineDown.EndPoint = new Point3d(cableLineUp.EndPoint.X, cableLine.EndPoint.Y, 0);
                    modSpace.AppendEntity(cableLineDown);
                    acTrans.AddNewlyCreatedDBObject(cableLineDown, true);

                    Point3d lastTermPoint = new Point3d();
                    curTermNumber = 1;
                    curPairNumber = 0;

                    Line tBoxJumper = new Line();
                    int pairNumUp = 0;
                    MText pairMark;
                    double lowestPoint = 0;
                    for (int k = 0; k < groups[i].Units.Count; k++)
                    {
                        if (groups[i].Units[k].param.Contains("TS") && groups[i].Units[k].equipTerminals.Count == 4)
                        {
                            #region withTS
                            curPairNumber++;
                            tbox.LastPairNumber++;

                            Line tBoxLineOut = new Line();
                            tBoxLineOut.SetDatabaseDefaults();
                            tBoxLineOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            tBoxLineOut.StartPoint = DrawTerminalBoxUnitTS(acTrans, modSpace, groups[i].Units[k].cableOutput[0].EndPoint, groups[i].Units[k], cableLineDown.EndPoint, tbox.LastPairNumber, curPairNumber, ref tbox);
                            tBoxLineOut.EndPoint = tBoxLineOut.StartPoint.Add(new Vector3d(0, -15, 0));
                            modSpace.AppendEntity(tBoxLineOut);
                            acTrans.AddNewlyCreatedDBObject(tBoxLineOut, true);

                            //lowestPoint = tBoxLineOut.EndPoint.Y;

                            Group group = groups[i];
                            Unit unit = groups[i].Units[k];
                            unit.equipPoint = tBoxLineOut.EndPoint;
                            groups[i].Units[k] = unit;

                            if (curPairNumber == 1)
                                curPairMark.Erase();

                            pairMark = new MText();
                            pairMark.SetDatabaseDefaults();
                            pairMark.TextHeight = 2.5;
                            //if (tst.Has("spds 2.5-0.85"))
                            //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                            pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            double X = groups[i].Units[k].cableOutput[0].EndPoint.X - 1;
                            double Y = groups[i].Units[k].cableOutput[0].EndPoint.Y + (groups[i].Units[k].cableOutput[0].StartPoint.Y - groups[i].Units[k].cableOutput[0].EndPoint.Y) / 2;
                            pairMark.Location = new Point3d(X, Y, 0);
                            pairMark.Contents = "PR" + (tbox.LastPairNumber < 10 ? "\\W" + widthFactor + ";0" + tbox.LastPairNumber.ToString() : "\\W" + widthFactor + ";" + tbox.LastPairNumber.ToString());
                            pairMark.Rotation = 1.5708;
                            pairMark.Attachment = AttachmentPoint.BottomCenter;
                            modSpace.AppendEntity(pairMark);
                            acTrans.AddNewlyCreatedDBObject(pairMark, true);

                            curPairNumber = 0;
                            curTermNumber = 1;
                            #endregion
                        }
                        else
                        {
                            #region withoutTS
                            if (groups[i].Units[k].cupboardName.Contains("VA"))
                            {
                                curPairNumber++;
                                pairNumUp++;

                                Point3d point = DrawTerminalBoxUnit(acTrans, modSpace, groups[i].Units[k].cableOutput[0].EndPoint, groups[i].Units[k], cableLineDown.EndPoint, ref tbox, pairNumUp, curPairNumber, out lastTermPoint);
                                if (groups[i].Units[k].param.ToLower() != "резерв")
                                {
                                    tBoxJumper = new Line();
                                    tBoxJumper.SetDatabaseDefaults();
                                    tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    tBoxJumper.StartPoint = point;
                                    tBoxJumper.EndPoint = point;
                                    modSpace.AppendEntity(tBoxJumper);
                                    acTrans.AddNewlyCreatedDBObject(tBoxJumper, true);
                                }
                                else
                                {
                                    tBoxJumper = new Line();
                                    tBoxJumper.SetDatabaseDefaults();
                                    tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    tBoxJumper.StartPoint = point;
                                    tBoxJumper.EndPoint = point;
                                }

                                pairMark = new MText();
                                pairMark.SetDatabaseDefaults();
                                pairMark.TextHeight = 2.5;
                                //if (tst.Has("spds 2.5-0.85"))
                                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                double X = groups[i].Units[k].cableOutput[0].EndPoint.X - 1;
                                double Y = groups[i].Units[k].cableOutput[0].EndPoint.Y + (groups[i].Units[k].cableOutput[0].StartPoint.Y - groups[i].Units[k].cableOutput[0].EndPoint.Y) / 2;
                                pairMark.Location = new Point3d(X, Y, 0);
                                pairMark.Contents = "PR" + (pairNumUp < 10 ? "\\W" + widthFactor + ";0" + pairNumUp.ToString() : "\\W" + widthFactor + ";" + pairNumUp.ToString());
                                pairMark.Rotation = 1.5708;
                                pairMark.Attachment = AttachmentPoint.BottomCenter;
                                modSpace.AppendEntity(pairMark);
                                acTrans.AddNewlyCreatedDBObject(pairMark, true);
                            }
                            else
                            {
                                curPairNumber++;
                                tbox.LastPairNumber++;

                                Point3d point = DrawTerminalBoxUnit(acTrans, modSpace, groups[i].Units[k].cableOutput[0].EndPoint, groups[i].Units[k], cableLineDown.EndPoint, ref tbox, tbox.LastPairNumber, curPairNumber, out lastTermPoint);
                                tBoxJumper = new Line();
                                tBoxJumper.SetDatabaseDefaults();
                                tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                tBoxJumper.StartPoint = point;
                                tBoxJumper.EndPoint = point;
                                modSpace.AppendEntity(tBoxJumper);
                                acTrans.AddNewlyCreatedDBObject(tBoxJumper, true);

                                pairMark = new MText();
                                pairMark.SetDatabaseDefaults();
                                pairMark.TextHeight = 2.5;
                                //if (tst.Has("spds 2.5-0.85"))
                                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                double X = groups[i].Units[k].cableOutput[0].EndPoint.X - 1;
                                double Y = groups[i].Units[k].cableOutput[0].EndPoint.Y + (groups[i].Units[k].cableOutput[0].StartPoint.Y - groups[i].Units[k].cableOutput[0].EndPoint.Y) / 2;
                                pairMark.Location = new Point3d(X, Y, 0);
                                pairMark.Contents = "PR" + (tbox.LastPairNumber < 10 ? "\\W" + widthFactor + ";0" + tbox.LastPairNumber.ToString() : "\\W" + widthFactor + ";" + tbox.LastPairNumber.ToString());
                                pairMark.Rotation = 1.5708;
                                pairMark.Attachment = AttachmentPoint.BottomCenter;
                                modSpace.AppendEntity(pairMark);
                                acTrans.AddNewlyCreatedDBObject(pairMark, true);
                            }

                            for (int j = 1; j < groups[i].Units[k].cableOutput.Count; j++)
                            {
                                if (groups[i].Units[k].cupboardName.Contains("VA"))
                                {
                                    curPairNumber++;
                                    pairNumUp++;

                                    tBoxJumper.EndPoint = DrawTerminalBoxUnit(acTrans, modSpace, groups[i].Units[k].cableOutput[j].EndPoint, groups[i].Units[k], cableLineDown.EndPoint, ref tbox, pairNumUp, curPairNumber, out lastTermPoint);

                                    pairMark = new MText();
                                    pairMark.SetDatabaseDefaults();
                                    pairMark.TextHeight = 2.5;
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                                    pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    double X = groups[i].Units[k].cableOutput[j].EndPoint.X - 1;
                                    double Y = groups[i].Units[k].cableOutput[j].EndPoint.Y + (groups[i].Units[k].cableOutput[j].StartPoint.Y - groups[i].Units[k].cableOutput[j].EndPoint.Y) / 2;
                                    pairMark.Location = new Point3d(X, Y, 0);
                                    pairMark.Contents = "PR" + (pairNumUp < 10 ? "\\W" + widthFactor + ";0" + pairNumUp.ToString() : "\\W" + widthFactor + ";" + pairNumUp.ToString());
                                    pairMark.Rotation = 1.5708;
                                    pairMark.Attachment = AttachmentPoint.BottomCenter;
                                    modSpace.AppendEntity(pairMark);
                                    acTrans.AddNewlyCreatedDBObject(pairMark, true);
                                }
                                else
                                {
                                    curPairNumber++;
                                    tbox.LastPairNumber++;

                                    tBoxJumper.EndPoint = DrawTerminalBoxUnit(acTrans, modSpace, groups[i].Units[k].cableOutput[j].EndPoint, groups[i].Units[k], cableLineDown.EndPoint, ref tbox, tbox.LastPairNumber, curPairNumber, out lastTermPoint);

                                    pairMark = new MText();
                                    pairMark.SetDatabaseDefaults();
                                    pairMark.TextHeight = 2.5;
                                    //if (tst.Has("spds 2.5-0.85"))
                                    //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                                    pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                    double X = groups[i].Units[k].cableOutput[j].EndPoint.X - 1;
                                    double Y = groups[i].Units[k].cableOutput[j].EndPoint.Y + (groups[i].Units[k].cableOutput[j].StartPoint.Y - groups[i].Units[k].cableOutput[j].EndPoint.Y) / 2;
                                    pairMark.Location = new Point3d(X, Y, 0);
                                    pairMark.Contents = "PR" + (tbox.LastPairNumber < 10 ? "\\W"+widthFactor+";0" + tbox.LastPairNumber.ToString() : "\\W" + widthFactor + ";" + tbox.LastPairNumber.ToString());
                                    pairMark.Rotation = 1.5708;
                                    pairMark.Attachment = AttachmentPoint.BottomCenter;
                                    modSpace.AppendEntity(pairMark);
                                    acTrans.AddNewlyCreatedDBObject(pairMark, true);
                                }
                            }

                            if (groups[i].Units[k].param.ToLower() != "резерв")
                            {
                                Line tBoxLineOut = new Line();
                                tBoxLineOut.SetDatabaseDefaults();
                                tBoxLineOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                tBoxLineOut.StartPoint = tBoxJumper.StartPoint + (tBoxJumper.EndPoint - tBoxJumper.StartPoint) / 2;
                                tBoxLineOut.EndPoint = tBoxLineOut.StartPoint.Add(new Vector3d(0, -15, 0));
                                modSpace.AppendEntity(tBoxLineOut);
                                acTrans.AddNewlyCreatedDBObject(tBoxLineOut, true);

                                Group group = groups[i];
                                Unit unit = groups[i].Units[k];
                                unit.equipPoint = tBoxLineOut.EndPoint;
                                groups[i].Units[k] = unit;

                                if (curPairNumber == 1)
                                    curPairMark.Erase();
                            }
                            else
                            {
                                Line tBoxLineOut = new Line();
                                tBoxLineOut.SetDatabaseDefaults();
                                tBoxLineOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                                tBoxLineOut.StartPoint = tBoxJumper.StartPoint + (tBoxJumper.EndPoint - tBoxJumper.StartPoint) / 2;
                                tBoxLineOut.EndPoint = tBoxLineOut.StartPoint.Add(new Vector3d(0, -15, 0));

                                Group group = groups[i];
                                Unit unit = groups[i].Units[k];
                                unit.equipPoint = tBoxLineOut.EndPoint;
                                groups[i].Units[k] = unit;
                            }
                            curPairNumber = 0;
                            curTermNumber = 1;
                            lowestPoint = tBoxJumper.StartPoint.Y;
                            #endregion
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
                    tBoxFrame.AddVertexAt(0, new Point2d(cableLineDown.StartPoint.X - (5 * groups[i].Units[groups[i].Units.Count - 1].terminals.Count) - 5, cableLineDown.EndPoint.Y + 7), 0, 0, 0);
                    tBoxFrame.AddVertexAt(1, new Point2d(cableLineDown.EndPoint.X + 14, cableLineDown.EndPoint.Y + 7), 0, 0, 0);
                    tBoxFrame.AddVertexAt(2, new Point2d(cableLineDown.EndPoint.X + 14, lowestPoint - 2), 0, 0, 0);
                    tBoxFrame.AddVertexAt(3, new Point2d(cableLineDown.StartPoint.X - (5 * groups[i].Units[groups[i].Units.Count - 1].terminals.Count) - 5, lowestPoint - 2), 0, 0, 0);
                    modSpace.AppendEntity(tBoxFrame);
                    acTrans.AddNewlyCreatedDBObject(tBoxFrame, true);

                    MText tBoxName = new MText();
                    tBoxName.SetDatabaseDefaults();
                    tBoxName.TextHeight = 3;
                    //if (tst.Has("spds 2.5-0.85"))
                    //    tBoxName.TextStyleId = tst["spds 2.5-0.85"];
                    tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    tBoxName.Location = tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(1, 6, 0));
                    tBoxName.Contents = "\\W" + widthFactor + ";" + groups[i].Units[groups[i].Units.Count - 1].tBoxName;
                    tBoxName.Attachment = AttachmentPoint.BottomLeft;
                    modSpace.AppendEntity(tBoxName);
                    acTrans.AddNewlyCreatedDBObject(tBoxName, true);
                    tbox.textName.Add(tBoxName);

                    Leader acLdr = new Leader();
                    acLdr.SetDatabaseDefaults();
                    acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(-5, 0, 0)));
                    acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(0, 5, 0)));
                    acLdr.AppendVertex(tBoxFrame.GetPoint3dAt(1).Add(new Vector3d(tBoxName.ActualWidth + 1, 5, 0)));
                    acLdr.HasArrowHead = false;
                    modSpace.AppendEntity(acLdr);
                    acTrans.AddNewlyCreatedDBObject(acLdr, true);
                    tbox.ldr.Add(acLdr);

                    if (groups[i].Units[groups[i].Units.Count - 1].shield)
                    {
                        tbox.shield = true;
                        tbox.lastTermPoint = lastTermPoint;
                        tbox.cableLineSPLX = cableLine.StartPoint.X - 3;
                        tbox.cableLineSPRX = cableLine.StartPoint.X + 3;
                        tbox.cableLineEPY = cableLine.EndPoint.Y + 8;
                        tbox.framePoint = tBoxFrame.GetPoint3dAt(2);
                    }
                    else tbox.shield = false;

                    tboxes[index] = tbox;
                }
                else
                {
                    cableLine.EndPoint = cableLine.EndPoint.Add(new Vector3d(0, -135, 0));
                    cableMark.Location = cableLine.EndPoint.Add(new Vector3d(-1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                    textName.Location = cableLine.EndPoint.Add(new Vector3d(1, (cableLine.StartPoint.Y - cableLine.EndPoint.Y) / 2, 0));
                    Unit unit = groups[i].Units[groups[i].Units.Count - 1];
                    unit.equipPoint = cableLine.EndPoint;
                    groups[i].Units[groups[i].Units.Count - 1] = unit;
                }
                for (int k = 0; k < groups[i].Units.Count; k++)
                {
                    if (groups[i].Units[k].newSheet) curTableNum++;
                    double length = (groups[i].Units[k].equipTerminals.Count - 2) * 5 + 5;
                    Line equipJumper = new Line();
                    if (groups[i].Units[k].param.ToLower() != "резерв")
                    {
                        equipJumper = new Line();
                        equipJumper.SetDatabaseDefaults();
                        equipJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        equipJumper.StartPoint = groups[i].Units[k].equipPoint.Add(new Vector3d(-length / 2, 0, 0));
                        equipJumper.EndPoint = equipJumper.StartPoint.Add(new Vector3d(length, 0, 0));
                        modSpace.AppendEntity(equipJumper);
                        acTrans.AddNewlyCreatedDBObject(equipJumper, true);

                        if (groups[i].Units[k].param.Contains("TT"))
                        {
                            MText star = new MText();
                            star.SetDatabaseDefaults();
                            star.TextHeight = 3;
                            //if (tst.Has("spds 2.5-0.85"))
                            //    star.TextStyleId = tst["spds 2.5-0.85"];
                            star.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            star.Location = groups[i].Units[k].equipPoint.Add(new Vector3d(-3, 10, 0));
                            star.Attachment = AttachmentPoint.TopCenter;
                            star.Rotation = 1.5708;
                            star.Contents = "*";
                            modSpace.AppendEntity(star);
                            acTrans.AddNewlyCreatedDBObject(star, true);
                        }
                    }
                    else
                    {
                        equipJumper = new Line();
                        equipJumper.SetDatabaseDefaults();
                        equipJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        equipJumper.StartPoint = groups[i].Units[k].equipPoint.Add(new Vector3d(-length / 2, 0, 0));
                        equipJumper.EndPoint = equipJumper.StartPoint.Add(new Vector3d(length, 0, 0));
                    }

                    double lowestPoint = 0;
                    int color = 0;
                    for (int c = 0; c < groups[i].Units[k].equipTerminals.Count; c++)
                    {
                        if (groups[i].Units[k].param.ToLower() != "резерв")
                        {
                            Line equipLine = new Line();
                            equipLine.SetDatabaseDefaults();
                            equipLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            equipLine.StartPoint = equipJumper.StartPoint.Add(new Vector3d(5 * c, 0, 0));
                            equipLine.EndPoint = equipLine.StartPoint.Add(new Vector3d(0, -10, 0));
                            modSpace.AppendEntity(equipLine);
                            acTrans.AddNewlyCreatedDBObject(equipLine, true);

                            MText colorMark = new MText();
                            colorMark.SetDatabaseDefaults();
                            colorMark.TextHeight = 2.5;
                            //if (tst.Has("spds 2.5-0.85"))
                            //    colorMark.TextStyleId = tst["spds 2.5-0.85"];
                            colorMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            double x = equipLine.EndPoint.X - 1;
                            double y = equipLine.EndPoint.Y + (equipLine.StartPoint.Y - equipLine.EndPoint.Y) / 2;
                            colorMark.Location = new Point3d(x, y, 0);
                            colorMark.Contents = "\\W" + widthFactor + ";" + groups[i].Units[k].colors[color];
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
                            //if (tst.Has("spds 2.5-0.85"))
                            //    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                            cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            x = equipTerm.GetPoint3dAt(0).X + (equipTerm.GetPoint3dAt(1).X - equipTerm.GetPoint3dAt(0).X) / 2;
                            y = equipTerm.GetPoint3dAt(2).Y + (equipTerm.GetPoint3dAt(1).Y - equipTerm.GetPoint3dAt(2).Y) / 2;
                            cableTextDown.Location = new Point3d(x, y, 0);
                            cableTextDown.Attachment = AttachmentPoint.MiddleCenter;
                            cableTextDown.Contents = "\\W" + widthFactor + ";" + groups[i].Units[k].equipTerminals[c];
                            modSpace.AppendEntity(cableTextDown);
                            acTrans.AddNewlyCreatedDBObject(cableTextDown, true);

                            lowestPoint = equipTerm.GetPoint3dAt(3).Y;
                            if (color < groups[i].Units[k].colors.Count - 1) color++;
                            else color = 0;
                        }
                        else
                        {
                            Line equipLine = new Line();
                            equipLine.SetDatabaseDefaults();
                            equipLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            equipLine.StartPoint = equipJumper.StartPoint.Add(new Vector3d(5 * c, 0, 0));
                            equipLine.EndPoint = equipLine.StartPoint.Add(new Vector3d(0, -10, 0));

                            MText colorMark = new MText();
                            colorMark.SetDatabaseDefaults();
                            colorMark.TextHeight = 2.5;
                            //if (tst.Has("spds 2.5-0.85"))
                            //    colorMark.TextStyleId = tst["spds 2.5-0.85"];
                            colorMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            double x = equipLine.EndPoint.X - 1;
                            double y = equipLine.EndPoint.Y + (equipLine.StartPoint.Y - equipLine.EndPoint.Y) / 2;
                            colorMark.Location = new Point3d(x, y, 0);
                            colorMark.Contents = "\\W" + widthFactor + ";" + groups[i].Units[k].colors[color];
                            colorMark.Rotation = 1.5708;
                            colorMark.Attachment = AttachmentPoint.BottomCenter;

                            Polyline equipTerm = new Polyline();
                            equipTerm.SetDatabaseDefaults();
                            equipTerm.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            equipTerm.Closed = true;
                            equipTerm.AddVertexAt(0, new Point2d(equipLine.EndPoint.X - 2.5, equipLine.EndPoint.Y), 0, 0, 0);
                            equipTerm.AddVertexAt(1, equipTerm.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                            equipTerm.AddVertexAt(2, equipTerm.GetPoint2dAt(1).Add(new Vector2d(0, -5)), 0, 0, 0);
                            equipTerm.AddVertexAt(3, equipTerm.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);

                            MText cableTextDown = new MText();
                            cableTextDown.SetDatabaseDefaults();
                            cableTextDown.TextHeight = 2.5;
                            //if (tst.Has("spds 2.5-0.85"))
                            //    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                            cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                            x = equipTerm.GetPoint3dAt(0).X + (equipTerm.GetPoint3dAt(1).X - equipTerm.GetPoint3dAt(0).X) / 2;
                            y = equipTerm.GetPoint3dAt(2).Y + (equipTerm.GetPoint3dAt(1).Y - equipTerm.GetPoint3dAt(2).Y) / 2;
                            cableTextDown.Location = new Point3d(x, y, 0);
                            cableTextDown.Attachment = AttachmentPoint.MiddleCenter;
                            cableTextDown.Contents = "\\W" + widthFactor + ";" + groups[i].Units[k].equipTerminals[c];

                            lowestPoint = equipTerm.GetPoint3dAt(3).Y;
                            if (color < groups[i].Units[k].colors.Count - 1) color++;
                            else color = 0;
                        }
                    }
                    
                    Polyline equipFrame = new Polyline();
                    if (groups[i].Units[k].param.ToLower() != "резерв")
                    {
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

                        //if (!groups[i].Units[k].designation.Contains("PDS") && !groups[i].Units[k].designation.Contains("MA") && !groups[i].Units[k].designation.Contains("MT"))
                        //{
                        //    Line line1 = new Line();
                        //    line1.SetDatabaseDefaults();
                        //    line1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        //    line1.StartPoint = equipFrame.GetPoint3dAt(2);
                        //    line1.EndPoint = line1.StartPoint.Add(new Vector3d(5, 0, 0));
                        //    modSpace.AppendEntity(line1);
                        //    acTrans.AddNewlyCreatedDBObject(line1, true);

                        //    Line line2 = new Line();
                        //    line2.SetDatabaseDefaults();
                        //    line2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        //    line2.StartPoint = line1.EndPoint;
                        //    line2.EndPoint = line2.StartPoint.Add(new Vector3d(0, -3, 0));
                        //    modSpace.AppendEntity(line2);
                        //    acTrans.AddNewlyCreatedDBObject(line2, true);

                        //    Line line3 = new Line();
                        //    line3.SetDatabaseDefaults();
                        //    line3.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        //    line3.StartPoint = line2.EndPoint.Add(new Vector3d(-3, 0, 0));
                        //    line3.EndPoint = line3.StartPoint.Add(new Vector3d(6, 0, 0));
                        //    modSpace.AppendEntity(line3);
                        //    acTrans.AddNewlyCreatedDBObject(line3, true);

                        //    Line line4 = new Line();
                        //    line4.SetDatabaseDefaults();
                        //    line4.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        //    line4.StartPoint = line2.EndPoint.Add(new Vector3d(-2, -1, 0));
                        //    line4.EndPoint = line4.StartPoint.Add(new Vector3d(4, 0, 0));
                        //    modSpace.AppendEntity(line4);
                        //    acTrans.AddNewlyCreatedDBObject(line4, true);

                        //    Line line5 = new Line();
                        //    line5.SetDatabaseDefaults();
                        //    line5.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        //    line5.StartPoint = line2.EndPoint.Add(new Vector3d(-1, -2, 0));
                        //    line5.EndPoint = line5.StartPoint.Add(new Vector3d(2, 0, 0));
                        //    modSpace.AppendEntity(line5);
                        //    acTrans.AddNewlyCreatedDBObject(line5, true);
                        //}
                    }
                    else
                    {

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
                    }

                    if (groups[i].Units[k].designation.Contains("FNM") && groups[i].Units[k].linkText!="")
                    {
                        MText tBoxName = new MText();
                        tBoxName.SetDatabaseDefaults();
                        tBoxName.TextHeight = 3;
                        //if (tst.Has("spds 2.5-0.85"))
                        //    tBoxName.TextStyleId = tst["spds 2.5-0.85"];
                        tBoxName.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxName.Location = equipFrame.GetPoint3dAt(1).Add(new Vector3d(6, -4, 0));
                        tBoxName.Contents = "\\W" + widthFactor + ";" + groups[i].Units[k].linkText;
                        tBoxName.Attachment = AttachmentPoint.BottomLeft;
                        modSpace.AppendEntity(tBoxName);
                        acTrans.AddNewlyCreatedDBObject(tBoxName, true);

                        Leader acLdr = new Leader();
                        acLdr.SetDatabaseDefaults();
                        acLdr.AppendVertex(equipFrame.GetPoint3dAt(1).Add(new Vector3d(0, -10, 0)));
                        acLdr.AppendVertex(equipFrame.GetPoint3dAt(1).Add(new Vector3d(5, -5, 0)));
                        acLdr.AppendVertex(equipFrame.GetPoint3dAt(1).Add(new Vector3d(5+tBoxName.ActualWidth + 2, -5, 0)));
                        acLdr.HasArrowHead = false;
                        modSpace.AppendEntity(acLdr);
                        acTrans.AddNewlyCreatedDBObject(acLdr, true);
                    }

                    if (rigthEdgeXTable==0)
                        rigthEdgeXTable = tables[curTableNum].Position.X + tables[curTableNum].Width;

                    double C = equipFrame.GetPoint2dAt(0).X;
                    double D = equipFrame.GetPoint2dAt(1).X;
                    double A = tables[curTableNum].Position.X;
                    double B = tables[curTableNum].Position.X + tables[curTableNum].Width;
                    width = 2 * (C - B) + (D - C);
                    tables[curTableNum].InsertColumns(tables[curTableNum].Columns.Count, width, 1);
                    tables[curTableNum].UnmergeCells(tables[curTableNum].Rows[0]);
                    //if (tst.Has("spds 2.5-0.85"))
                    //    tables[curTableNum].Columns[tables[curTableNum].Columns.Count - 1].TextStyleId = tst["spds 2.5-0.85"];
                    tables[curTableNum].Columns[tables[curTableNum].Columns.Count - 1].TextHeight = 2.5;
                    tables[curTableNum].Cells[0, tables[curTableNum].Columns.Count - 1].TextString = groups[i].Units[k].equipType == "" ? "-" : groups[i].Units[k].equipType;
                    tables[curTableNum].Cells[1, tables[curTableNum].Columns.Count - 1].TextString = groups[i].Units[k].designation;
                    tables[curTableNum].Cells[2, tables[curTableNum].Columns.Count - 1].TextString = groups[i].Units[k].param;
                    tables[curTableNum].Cells[3, tables[curTableNum].Columns.Count - 1].TextString = groups[i].Units[k].equipment;
                    tables[curTableNum].GenerateLayout();
                }
            }
        }

        private static Point3d DrawTerminalBoxUnit(Transaction acTrans, BlockTableRecord modSpace, Point3d point, Unit unit, Point3d cableEndPoint, ref tBox tbox, int pairNumberUp, int pairNumberDown, out Point3d lastTerminalPoint)
        {
            Line tBoxInput = new Line();
            tBoxInput.SetDatabaseDefaults();
            tBoxInput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInput.StartPoint = new Point3d(point.X, cableEndPoint.Y, 0);
            tBoxInput.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(tBoxInput);
            acTrans.AddNewlyCreatedDBObject(tBoxInput, true);

            MText pairMark = new MText();
            pairMark.SetDatabaseDefaults();
            pairMark.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
            pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double X = tBoxInput.EndPoint.X - 1;
            double Y = tBoxInput.EndPoint.Y + (tBoxInput.StartPoint.Y - tBoxInput.EndPoint.Y) / 2;
            pairMark.Location = new Point3d(X, Y, 0);
            pairMark.Contents = "PR" + (pairNumberUp < 10 ? "\\W" + widthFactor + ";0" + pairNumberUp.ToString() : "\\W" + widthFactor + ";" + pairNumberUp.ToString());
            pairMark.Rotation = 1.5708;
            pairMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(pairMark);
            acTrans.AddNewlyCreatedDBObject(pairMark, true);

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
            for (int i = unit.colors.Count - 1; i >= 0; i--)
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
                //if (tst.Has("spds 2.5-0.85"))
                //    cableTextUp.TextStyleId = tst["spds 2.5-0.85"];
                cableTextUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                double x = cablePolyUp.GetPoint3dAt(0).X + (cablePolyUp.GetPoint3dAt(1).X - cablePolyUp.GetPoint3dAt(0).X) / 2;
                double y = cablePolyUp.GetPoint3dAt(2).Y + (cablePolyUp.GetPoint3dAt(1).Y - cablePolyUp.GetPoint3dAt(2).Y) / 2;
                cableTextUp.Location = new Point3d(x, y, 0);
                cableTextUp.Contents = unit.param.ToLower() != "резерв" ? "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber : "\\W"+widthFactor+";Резерв";
                cableTextUp.Rotation = 1.5708;
                cableTextUp.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(cableTextUp);
                acTrans.AddNewlyCreatedDBObject(cableTextUp, true);

                Line colorLineUp = new Line();
                colorLineUp.SetDatabaseDefaults();
                colorLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                colorLineUp.StartPoint = cablePolyUp.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                colorLineUp.EndPoint = colorLineUp.StartPoint.Add(new Vector3d(0, -10, 0));
                modSpace.AppendEntity(colorLineUp);
                acTrans.AddNewlyCreatedDBObject(colorLineUp, true);

                MText colorMarkUp = new MText();
                colorMarkUp.SetDatabaseDefaults();
                colorMarkUp.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    colorMarkUp.TextStyleId = tst["spds 2.5-0.85"];
                colorMarkUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = colorLineUp.EndPoint.X - 1;
                y = colorLineUp.EndPoint.Y + (colorLineUp.StartPoint.Y - colorLineUp.EndPoint.Y) / 2;
                colorMarkUp.Location = new Point3d(x, y, 0);
                colorMarkUp.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
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
                //if (tst.Has("spds 2.5-0.85"))
                //    termText.TextStyleId = tst["spds 2.5-0.85"];
                termText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2;
                y = termPoly.GetPoint3dAt(2).Y + (termPoly.GetPoint3dAt(1).Y - termPoly.GetPoint3dAt(2).Y) / 2;
                termText.Location = new Point3d(x, y, 0);
                termText.Contents = "\\W" + widthFactor + ";" + tbox.LastTerminalNumber.ToString();
                termText.Rotation = 1.5708;
                termText.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(termText);
                acTrans.AddNewlyCreatedDBObject(termText, true);
                tbox.LastTerminalNumber++;

                if (unit.param.ToLower() != "резерв")
                {
                    Line colorLineDown = new Line();
                    colorLineDown.SetDatabaseDefaults();
                    colorLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    colorLineDown.StartPoint = termPoly.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                    colorLineDown.EndPoint = colorLineDown.StartPoint.Add(new Vector3d(0, -10, 0));
                    modSpace.AppendEntity(colorLineDown);
                    acTrans.AddNewlyCreatedDBObject(colorLineDown, true);

                    MText colorMarkDown = new MText();
                    colorMarkDown.SetDatabaseDefaults();
                    colorMarkDown.TextHeight = 2.5;
                    //if (tst.Has("spds 2.5-0.85"))
                        //colorMarkDown.TextStyleId = tst["spds 2.5-0.85"];
                    colorMarkDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    x = colorLineDown.EndPoint.X - 1;
                    y = colorLineDown.EndPoint.Y + (colorLineDown.StartPoint.Y - colorLineDown.EndPoint.Y) / 2;
                    colorMarkDown.Location = new Point3d(x, y, 0);
                    colorMarkDown.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
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
                    //if (tst.Has("spds 2.5-0.85"))
                    //    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                    cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    x = cablePolyDown.GetPoint3dAt(0).X + (cablePolyDown.GetPoint3dAt(1).X - cablePolyDown.GetPoint3dAt(0).X) / 2;
                    y = cablePolyDown.GetPoint3dAt(2).Y + (cablePolyDown.GetPoint3dAt(1).Y - cablePolyDown.GetPoint3dAt(2).Y) / 2;
                    cableTextDown.Location = new Point3d(x, y, 0);
                    cableTextDown.Contents = "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber;
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
                }
                else
                {
                    Line colorLineDown = new Line();
                    colorLineDown.SetDatabaseDefaults();
                    colorLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    colorLineDown.StartPoint = termPoly.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                    colorLineDown.EndPoint = colorLineDown.StartPoint.Add(new Vector3d(0, -10, 0));

                    MText colorMarkDown = new MText();
                    colorMarkDown.SetDatabaseDefaults();
                    colorMarkDown.TextHeight = 2.5;
                    //if (tst.Has("spds 2.5-0.85"))
                    //    colorMarkDown.TextStyleId = tst["spds 2.5-0.85"];
                    colorMarkDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    x = colorLineDown.EndPoint.X - 1;
                    y = colorLineDown.EndPoint.Y + (colorLineDown.StartPoint.Y - colorLineDown.EndPoint.Y) / 2;
                    colorMarkDown.Location = new Point3d(x, y, 0);
                    colorMarkDown.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
                    colorMarkDown.Rotation = 1.5708;
                    colorMarkDown.Attachment = AttachmentPoint.BottomCenter;

                    Polyline cablePolyDown = new Polyline();
                    cablePolyDown.SetDatabaseDefaults();
                    cablePolyDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cablePolyDown.Closed = true;
                    cablePolyDown.AddVertexAt(0, new Point2d(colorLineDown.EndPoint.X - 2.5, colorLineDown.EndPoint.Y), 0, 0, 0);
                    cablePolyDown.AddVertexAt(1, cablePolyDown.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                    cablePolyDown.AddVertexAt(2, cablePolyDown.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
                    cablePolyDown.AddVertexAt(3, cablePolyDown.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);

                    MText cableTextDown = new MText();
                    cableTextDown.SetDatabaseDefaults();
                    cableTextDown.TextHeight = 2.5;
                    //if (tst.Has("spds 2.5-0.85"))
                    //    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                    cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    x = cablePolyDown.GetPoint3dAt(0).X + (cablePolyDown.GetPoint3dAt(1).X - cablePolyDown.GetPoint3dAt(0).X) / 2;
                    y = cablePolyDown.GetPoint3dAt(2).Y + (cablePolyDown.GetPoint3dAt(1).Y - cablePolyDown.GetPoint3dAt(2).Y) / 2;
                    cableTextDown.Location = new Point3d(x, y, 0);
                    cableTextDown.Contents = "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber;
                    cableTextDown.Rotation = 1.5708;
                    cableTextDown.Attachment = AttachmentPoint.MiddleCenter;

                    Line tBoxBranchOutput = new Line();
                    tBoxBranchOutput.SetDatabaseDefaults();
                    tBoxBranchOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    tBoxBranchOutput.StartPoint = cablePolyDown.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                    tBoxBranchOutput.EndPoint = tBoxBranchOutput.StartPoint.Add(new Vector3d(0, -8, 0));

                    lowestPoint = tBoxBranchOutput.EndPoint.Y;
                }
                curTermNumber++;
                if (color < unit.colors.Count - 1) color++;
                else color = 0;
            }
           
            Line tBoxOutputBranch = new Line();
            tBoxInputBranch.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(-5 * (unit.colors.Count - 1), -8, 0));
            if (unit.param.ToLower() != "резерв")
            {
                tBoxOutputBranch.SetDatabaseDefaults();
                tBoxOutputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutputBranch.StartPoint = new Point3d(tBoxInputBranch.StartPoint.X, lowestPoint, 0);
                tBoxOutputBranch.EndPoint = tBoxOutputBranch.StartPoint.Add(new Vector3d(-5 * (unit.colors.Count - 1), 0, 0));
                modSpace.AppendEntity(tBoxOutputBranch);
                acTrans.AddNewlyCreatedDBObject(tBoxOutputBranch, true);
            }
            else
            {
                tBoxOutputBranch.SetDatabaseDefaults();
                tBoxOutputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutputBranch.StartPoint = new Point3d(tBoxInputBranch.StartPoint.X, lowestPoint, 0);
                tBoxOutputBranch.EndPoint = tBoxOutputBranch.StartPoint.Add(new Vector3d(-5 * (unit.colors.Count - 1), 0, 0));
            }

            if (unit.shield && unit.cableMark.Contains("ИЭ"))
            {
                if (unit.param.ToLower() != "резерв")
                {
                    lastTerminalPoint = DrawTBoxGnd(acTrans, modSpace, tbox, lastTermPoint, tBoxInputBranch.EndPoint.X, tBoxInputBranch.StartPoint.X, tBoxInputBranch.StartPoint.Y, tBoxOutputBranch.StartPoint.Y, false, new Point3d());
                    tbox.LastShieldNumber++;
                }
                else
                {
                    lastTerminalPoint = DrawTBoxGnd(acTrans, modSpace, tbox, lastTermPoint, tBoxInputBranch.EndPoint.X, tBoxInputBranch.StartPoint.X, tBoxInputBranch.StartPoint.Y, tBoxInputBranch.StartPoint.Y, false, new Point3d());
                    tbox.LastShieldNumber++;
                }
            }
            else lastTerminalPoint = lastTermPoint;
            
            if (unit.param.ToLower() != "резерв")
            {
                Line tBoxOutput = new Line();
                tBoxOutput.SetDatabaseDefaults();
                tBoxOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutput.StartPoint = tBoxOutputBranch.StartPoint;
                tBoxOutput.EndPoint = tBoxOutput.StartPoint.Add(new Vector3d(0, -10, 0));
                modSpace.AppendEntity(tBoxOutput);
                acTrans.AddNewlyCreatedDBObject(tBoxOutput, true);

                pairMark = new MText();
                curPairMark = pairMark;
                pairMark.SetDatabaseDefaults();
                pairMark.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                X = tBoxOutput.EndPoint.X - 1;
                Y = tBoxOutput.EndPoint.Y + (tBoxOutput.StartPoint.Y - tBoxOutput.EndPoint.Y) / 2;
                pairMark.Location = new Point3d(X, Y, 0);
                pairMark.Contents = "PR" + (curPairNumber < 10 ? "\\W" + widthFactor + ";0" + pairNumberDown.ToString() : "\\W" + widthFactor + ";" + pairNumberDown.ToString());
                pairMark.Rotation = 1.5708;
                pairMark.Attachment = AttachmentPoint.BottomCenter;
                modSpace.AppendEntity(pairMark);
                acTrans.AddNewlyCreatedDBObject(pairMark, true);

                return tBoxOutput.EndPoint;
            } 
            else
            {
                Line tBoxOutput = new Line();
                tBoxOutput.SetDatabaseDefaults();
                tBoxOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutput.StartPoint = tBoxOutputBranch.StartPoint;
                tBoxOutput.EndPoint = tBoxOutput.StartPoint.Add(new Vector3d(0, -10, 0));

                pairMark = new MText();
                curPairMark = pairMark;
                pairMark.SetDatabaseDefaults();
                pairMark.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                X = tBoxOutput.EndPoint.X - 1;
                Y = tBoxOutput.EndPoint.Y + (tBoxOutput.StartPoint.Y - tBoxOutput.EndPoint.Y) / 2;
                pairMark.Location = new Point3d(X, Y, 0);
                pairMark.Contents = "PR" + (curPairNumber < 10 ? "\\W" + widthFactor + ";0" + pairNumberDown.ToString() : "\\W" + widthFactor + ";" + pairNumberDown.ToString());
                pairMark.Rotation = 1.5708;
                pairMark.Attachment = AttachmentPoint.BottomCenter;

                return tBoxOutput.EndPoint;
            }
        }
        private static Point3d DrawTerminalBoxUnitTS(Transaction acTrans, BlockTableRecord modSpace, Point3d point, Unit unit, Point3d cableEndPoint, int pairNumberUp, int pairNumberDown, ref tBox tbox)
        {
            Line tBoxInput = new Line();
            tBoxInput.SetDatabaseDefaults();
            tBoxInput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInput.StartPoint = new Point3d(point.X, cableEndPoint.Y, 0);
            tBoxInput.EndPoint = tBoxInput.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(tBoxInput);
            acTrans.AddNewlyCreatedDBObject(tBoxInput, true);

            MText pairMark = new MText();
            pairMark.SetDatabaseDefaults();
            pairMark.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
            pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double X = tBoxInput.EndPoint.X - 1;
            double Y = tBoxInput.EndPoint.Y + (tBoxInput.StartPoint.Y - tBoxInput.EndPoint.Y) / 2;
            pairMark.Location = new Point3d(X, Y, 0);
            pairMark.Contents = "PR" + (pairNumberUp < 10 ? "\\W" + widthFactor + ";0" + pairNumberUp.ToString() : "\\W" + widthFactor + ";" + pairNumberUp.ToString());
            pairMark.Rotation = 1.5708;
            pairMark.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(pairMark);
            acTrans.AddNewlyCreatedDBObject(pairMark, true);

            Line tBoxInputBranch = new Line();
            tBoxInputBranch.SetDatabaseDefaults();
            tBoxInputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInputBranch.StartPoint = tBoxInput.EndPoint.Add(new Vector3d(-7.5, 0, 0));
            tBoxInputBranch.EndPoint = tBoxInputBranch.StartPoint.Add(new Vector3d(15, 0, 0));
            modSpace.AppendEntity(tBoxInputBranch);
            acTrans.AddNewlyCreatedDBObject(tBoxInputBranch, true);

            double lowestPoint = 0;
            Point3d lastTermPoint = new Point3d();
            int color = 0;

            Line tBoxInputBranchDown = new Line();
            tBoxInputBranchDown.SetDatabaseDefaults();
            tBoxInputBranchDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInputBranchDown.StartPoint = tBoxInputBranch.StartPoint;
            tBoxInputBranchDown.EndPoint = tBoxInputBranchDown.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(tBoxInputBranchDown);
            acTrans.AddNewlyCreatedDBObject(tBoxInputBranchDown, true);

            Polyline cablePolyUp = new Polyline();
            cablePolyUp.SetDatabaseDefaults();
            cablePolyUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cablePolyUp.Closed = true;
            cablePolyUp.AddVertexAt(0, new Point2d(tBoxInputBranchDown.EndPoint.X - 2.5, tBoxInputBranchDown.EndPoint.Y), 0, 0, 0);
            cablePolyUp.AddVertexAt(1, cablePolyUp.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
            cablePolyUp.AddVertexAt(2, cablePolyUp.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
            cablePolyUp.AddVertexAt(3, cablePolyUp.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
            modSpace.AppendEntity(cablePolyUp);
            acTrans.AddNewlyCreatedDBObject(cablePolyUp, true);
            point = point.Add(new Vector3d(5, 0, 0));

            MText cableTextUp = new MText();
            cableTextUp.SetDatabaseDefaults();
            cableTextUp.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    cableTextUp.TextStyleId = tst["spds 2.5-0.85"];
            cableTextUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = cablePolyUp.GetPoint3dAt(0).X + (cablePolyUp.GetPoint3dAt(1).X - cablePolyUp.GetPoint3dAt(0).X) / 2;
            double y = cablePolyUp.GetPoint3dAt(2).Y + (cablePolyUp.GetPoint3dAt(1).Y - cablePolyUp.GetPoint3dAt(2).Y) / 2;
            cableTextUp.Location = new Point3d(x, y, 0);
            cableTextUp.Contents = unit.param.ToLower() != "резерв" ? "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber : "\\W"+widthFactor+";Резерв";
            cableTextUp.Rotation = 1.5708;
            cableTextUp.Attachment = AttachmentPoint.MiddleCenter;
            modSpace.AppendEntity(cableTextUp);
            acTrans.AddNewlyCreatedDBObject(cableTextUp, true);

            Line colorLineUp = new Line();
            colorLineUp.SetDatabaseDefaults();
            colorLineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            colorLineUp.StartPoint = cablePolyUp.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
            colorLineUp.EndPoint = colorLineUp.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(colorLineUp);
            acTrans.AddNewlyCreatedDBObject(colorLineUp, true);

            MText colorMarkUp = new MText();
            colorMarkUp.SetDatabaseDefaults();
            colorMarkUp.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    colorMarkUp.TextStyleId = tst["spds 2.5-0.85"];
            colorMarkUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            x = colorLineUp.EndPoint.X - 1;
            y = colorLineUp.EndPoint.Y + (colorLineUp.StartPoint.Y - colorLineUp.EndPoint.Y) / 2;
            colorMarkUp.Location = new Point3d(x, y, 0);
            colorMarkUp.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
            colorMarkUp.Rotation = 1.5708;
            colorMarkUp.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(colorMarkUp);
            acTrans.AddNewlyCreatedDBObject(colorMarkUp, true);

            lastTermPoint = new Point3d(colorLineUp.EndPoint.X - 2.5, colorLineUp.EndPoint.Y, 0);
            double[] edges = new double[4];
            for (int i = 0; i < 4; i++)
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
                lastTermPoint = termPoly.GetPoint3dAt(1);

                MText termText = new MText();
                termText.SetDatabaseDefaults();
                termText.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    termText.TextStyleId = tst["spds 2.5-0.85"];
                termText.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = termPoly.GetPoint3dAt(0).X + (termPoly.GetPoint3dAt(1).X - termPoly.GetPoint3dAt(0).X) / 2;
                y = termPoly.GetPoint3dAt(2).Y + (termPoly.GetPoint3dAt(1).Y - termPoly.GetPoint3dAt(2).Y) / 2;
                termText.Location = new Point3d(x, y, 0);
                termText.Contents = "\\W" + widthFactor + ";" + tbox.LastTerminalNumber.ToString();
                termText.Rotation = 1.5708;
                termText.Attachment = AttachmentPoint.MiddleCenter;
                modSpace.AppendEntity(termText);
                acTrans.AddNewlyCreatedDBObject(termText, true);
                tbox.LastTerminalNumber++;

                Line colorLineDown = new Line();
                colorLineDown.SetDatabaseDefaults();
                colorLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                colorLineDown.StartPoint = termPoly.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
                colorLineDown.EndPoint = colorLineDown.StartPoint.Add(new Vector3d(0, -10, 0));
                modSpace.AppendEntity(colorLineDown);
                acTrans.AddNewlyCreatedDBObject(colorLineDown, true);

                MText colorMarkDown = new MText();
                colorMarkDown.SetDatabaseDefaults();
                colorMarkDown.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    colorMarkDown.TextStyleId = tst["spds 2.5-0.85"];
                colorMarkDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = colorLineDown.EndPoint.X - 1;
                y = colorLineDown.EndPoint.Y + (colorLineDown.StartPoint.Y - colorLineDown.EndPoint.Y) / 2;
                colorMarkDown.Location = new Point3d(x, y, 0);
                colorMarkDown.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
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
                //if (tst.Has("spds 2.5-0.85"))
                //    cableTextDown.TextStyleId = tst["spds 2.5-0.85"];
                cableTextDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                x = cablePolyDown.GetPoint3dAt(0).X + (cablePolyDown.GetPoint3dAt(1).X - cablePolyDown.GetPoint3dAt(0).X) / 2;
                y = cablePolyDown.GetPoint3dAt(2).Y + (cablePolyDown.GetPoint3dAt(1).Y - cablePolyDown.GetPoint3dAt(2).Y) / 2;
                cableTextDown.Location = new Point3d(x, y, 0);
                cableTextDown.Contents = "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber;
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
                edges[i] = tBoxBranchOutput.EndPoint.X;

                curTermNumber++;
                if (color < unit.colors.Count - 1) color++;
                else color = 0;
            }
            if (color == 0) color = unit.colors.Count - 1;
            else color--;
            curTermNumber--;

            Line tBoxInputBranchDown1 = new Line();
            tBoxInputBranchDown1.SetDatabaseDefaults();
            tBoxInputBranchDown1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxInputBranchDown1.StartPoint = tBoxInputBranch.EndPoint;
            tBoxInputBranchDown1.EndPoint = tBoxInputBranchDown1.StartPoint.Add(new Vector3d(0, -8, 0));
            modSpace.AppendEntity(tBoxInputBranchDown1);
            acTrans.AddNewlyCreatedDBObject(tBoxInputBranchDown1, true);

            Polyline cablePolyUp1 = new Polyline();
            cablePolyUp1.SetDatabaseDefaults();
            cablePolyUp1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cablePolyUp1.Closed = true;
            cablePolyUp1.AddVertexAt(0, new Point2d(tBoxInputBranchDown1.EndPoint.X - 2.5, tBoxInputBranchDown1.EndPoint.Y), 0, 0, 0);
            cablePolyUp1.AddVertexAt(1, cablePolyUp1.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
            cablePolyUp1.AddVertexAt(2, cablePolyUp1.GetPoint2dAt(1).Add(new Vector2d(0, -30)), 0, 0, 0);
            cablePolyUp1.AddVertexAt(3, cablePolyUp1.GetPoint2dAt(0).Add(new Vector2d(0, -30)), 0, 0, 0);
            modSpace.AppendEntity(cablePolyUp1);
            acTrans.AddNewlyCreatedDBObject(cablePolyUp1, true);
            point = point.Add(new Vector3d(5, 0, 0));

            MText cableTextUp1 = new MText();
            cableTextUp1.SetDatabaseDefaults();
            cableTextUp1.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    cableTextUp1.TextStyleId = tst["spds 2.5-0.85"];
            cableTextUp1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            x = cablePolyUp1.GetPoint3dAt(0).X + (cablePolyUp1.GetPoint3dAt(1).X - cablePolyUp1.GetPoint3dAt(0).X) / 2;
            y = cablePolyUp1.GetPoint3dAt(2).Y + (cablePolyUp1.GetPoint3dAt(1).Y - cablePolyUp1.GetPoint3dAt(2).Y) / 2;
            cableTextUp1.Location = new Point3d(x, y, 0);
            cableTextUp1.Contents = unit.param.ToLower() != "резерв" ? "\\W" + widthFactor + ";" + unit.designation + "/" + curTermNumber : "\\W"+widthFactor+";Резерв";
            cableTextUp1.Rotation = 1.5708;
            cableTextUp1.Attachment = AttachmentPoint.MiddleCenter;
            modSpace.AppendEntity(cableTextUp1);
            acTrans.AddNewlyCreatedDBObject(cableTextUp1, true);

            Line colorLineUp1 = new Line();
            colorLineUp1.SetDatabaseDefaults();
            colorLineUp1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            colorLineUp1.StartPoint = cablePolyUp1.GetPoint3dAt(2).Add(new Vector3d(-2.5, 0, 0));
            colorLineUp1.EndPoint = colorLineUp1.StartPoint.Add(new Vector3d(0, -10, 0));
            modSpace.AppendEntity(colorLineUp1);
            acTrans.AddNewlyCreatedDBObject(colorLineUp1, true);

            MText colorMarkUp1 = new MText();
            colorMarkUp1.SetDatabaseDefaults();
            colorMarkUp1.TextHeight = 2.5;
            //if (tst.Has("spds 2.5-0.85"))
            //    colorMarkUp1.TextStyleId = tst["spds 2.5-0.85"];
            colorMarkUp1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            x = colorLineUp1.EndPoint.X - 1;
            y = colorLineUp1.EndPoint.Y + (colorLineUp1.StartPoint.Y - colorLineUp1.EndPoint.Y) / 2;
            colorMarkUp1.Location = new Point3d(x, y, 0);
            colorMarkUp1.Contents = "\\W" + widthFactor + ";" + unit.colors[color];
            colorMarkUp1.Rotation = 1.5708;
            colorMarkUp1.Attachment = AttachmentPoint.BottomCenter;
            modSpace.AppendEntity(colorMarkUp1);
            acTrans.AddNewlyCreatedDBObject(colorMarkUp1, true);

            Line line1 = new Line();
            line1 = new Line();
            line1.SetDatabaseDefaults();
            line1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            line1.StartPoint = colorLineUp1.EndPoint.Add(new Vector3d(-5, 0, 0));
            line1.EndPoint = line1.StartPoint.Add(new Vector3d(0, 5, 0));
            modSpace.AppendEntity(line1);
            acTrans.AddNewlyCreatedDBObject(line1, true);

            Line line2 = new Line();
            line2 = new Line();
            line2.SetDatabaseDefaults();
            line2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            line2.StartPoint = line1.StartPoint.Add(new Vector3d(-5, 0, 0));
            line2.EndPoint = line1.EndPoint.Add(new Vector3d(-5, 0, 0));
            modSpace.AppendEntity(line2);
            acTrans.AddNewlyCreatedDBObject(line2, true);

            Line tBoxJumper = new Line();
            tBoxJumper.SetDatabaseDefaults();
            tBoxJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxJumper.StartPoint = line1.EndPoint;
            tBoxJumper.EndPoint = line2.EndPoint;
            modSpace.AppendEntity(tBoxJumper);
            acTrans.AddNewlyCreatedDBObject(tBoxJumper, true);

            Line tBoxOutputBranch = new Line();
            tBoxOutputBranch.SetDatabaseDefaults();
            tBoxOutputBranch.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxOutputBranch.StartPoint = new Point3d(edges[0], lowestPoint, 0);
            tBoxOutputBranch.EndPoint = new Point3d(edges[3], lowestPoint, 0);
            modSpace.AppendEntity(tBoxOutputBranch);
            acTrans.AddNewlyCreatedDBObject(tBoxOutputBranch, true);

            if (unit.shield && unit.cableMark.Contains("ИЭ"))
            {
                if (unit.param.ToLower() != "резерв")
                {
                    lastTermPoint = DrawTBoxGnd(acTrans, modSpace, tbox, lastTermPoint, tBoxInputBranch.StartPoint.X, tBoxInputBranch.EndPoint.X, tBoxInputBranch.StartPoint.Y, tBoxOutputBranch.StartPoint.Y, false, new Point3d());
                    tbox.LastShieldNumber++;
                }
                else
                {
                    lastTermPoint = DrawTBoxGnd(acTrans, modSpace, tbox, lastTermPoint, tBoxInputBranch.EndPoint.X, tBoxInputBranch.StartPoint.X, tBoxInputBranch.StartPoint.Y, tBoxInputBranch.StartPoint.Y, false, new Point3d());
                    tbox.LastShieldNumber++;
                }
            }
            //else lastTermPoint = tBoxInput.EndPoint.Add(new Vector3d(5, 0, 0));

            if (unit.param.ToLower() != "резерв")
            {
                Line tBoxOutput = new Line();
                tBoxOutput.SetDatabaseDefaults();
                tBoxOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutput.StartPoint = tBoxOutputBranch.StartPoint + (tBoxOutputBranch.EndPoint - tBoxOutputBranch.StartPoint)/2;
                tBoxOutput.EndPoint = tBoxOutput.StartPoint.Add(new Vector3d(0, -10, 0));
                modSpace.AppendEntity(tBoxOutput);
                acTrans.AddNewlyCreatedDBObject(tBoxOutput, true);

                pairMark = new MText();
                curPairMark = pairMark;
                pairMark.SetDatabaseDefaults();
                pairMark.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                X = tBoxOutput.EndPoint.X - 1;
                Y = tBoxOutput.EndPoint.Y + (tBoxOutput.StartPoint.Y - tBoxOutput.EndPoint.Y) / 2;
                pairMark.Location = new Point3d(X, Y, 0);
                pairMark.Contents = "PR" + (curPairNumber < 10 ? "\\W" + widthFactor + ";0" + pairNumberDown.ToString() : "\\W" + widthFactor + ";" + pairNumberDown.ToString());
                pairMark.Rotation = 1.5708;
                pairMark.Attachment = AttachmentPoint.BottomCenter;
                modSpace.AppendEntity(pairMark);
                acTrans.AddNewlyCreatedDBObject(pairMark, true);

                return tBoxOutput.EndPoint;
            }
            else
            {
                Line tBoxOutput = new Line();
                tBoxOutput.SetDatabaseDefaults();
                tBoxOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                tBoxOutput.StartPoint = tBoxOutputBranch.StartPoint;
                tBoxOutput.EndPoint = tBoxOutput.StartPoint.Add(new Vector3d(0, -10, 0));

                pairMark = new MText();
                curPairMark = pairMark;
                pairMark.SetDatabaseDefaults();
                pairMark.TextHeight = 2.5;
                //if (tst.Has("spds 2.5-0.85"))
                //    pairMark.TextStyleId = tst["spds 2.5-0.85"];
                pairMark.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                X = tBoxOutput.EndPoint.X - 1;
                Y = tBoxOutput.EndPoint.Y + (tBoxOutput.StartPoint.Y - tBoxOutput.EndPoint.Y) / 2;
                pairMark.Location = new Point3d(X, Y, 0);
                pairMark.Contents = "PR" + (curPairNumber < 10 ? "\\W" + widthFactor + ";0" + pairNumberDown.ToString() : "\\W" + widthFactor + ";" + pairNumberDown.ToString());
                pairMark.Rotation = 1.5708;
                pairMark.Attachment = AttachmentPoint.BottomCenter;

                return tBoxOutput.EndPoint;
            }
        }

        private static void InsertSheet(Transaction acTrans, BlockTableRecord modSpace, Database acdb, Point3d point)
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

        private static void LoadFonts(Transaction acTrans, BlockTableRecord modSpace, Database acdb)
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
