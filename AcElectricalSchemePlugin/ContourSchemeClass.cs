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
    static class ContourSchemeClass
    {
        private static List<UnitA> UnitsA;
        private static List<UnitD> UnitsD;
        private static List<Unit?> Units;
        private static List<CCUnit?> CCUnits;
        private static Document acDoc;
        private static Editor editor;
        private static List<Bus> busXT;
        private static List<Bus> busXR;
        private static List<Group> Groups;
        private static TextStyleTable tst;
        private static int CupboardNum;
        private static LinetypeTable lineTypeTable;

        struct UnitA
        {
            public string Equipment;
            public string Article;
            public int Rack;
            public int Slot;
            public int Channel;
            public string Type;
            public List<string> Fu;
            public List<string> Pins;
            public string Terminals;
            public List<Point3d> PinPoint;
        }

        struct Bus
        {
            public List<BusPoint> BusPoints;
            public string Term;
            public string Link;
            public string Type;

            public void SetPageLastBusPoint(bool page)
            {
                if (BusPoints!=null)
                {
                    BusPoint bP = BusPoints[BusPoints.Count - 1];
                    bP.Page = page;
                    BusPoints[BusPoints.Count - 1] = bP;
                }
            }

            public void SetEndLastBusPoint(bool end)
            {
                if (BusPoints != null)
                {
                    BusPoint bP = BusPoints[BusPoints.Count - 1];
                    bP.End = end;
                    BusPoints[BusPoints.Count - 1] = bP;
                }
            }

            public void AddPointLastBusPoint(Point3d point)
            {
                if (BusPoints != null)
                {
                    BusPoint bP = new BusPoint();
                    bP.Point = point;
                    bP.End = true;
                    BusPoints.Add(bP);
                }
            }

            public void SetPinFirstBusPoint(int pin)
            {
                if (BusPoints != null)
                {
                    BusPoint bP = BusPoints[0];
                    bP.Pin = pin;
                    BusPoints[0] = bP;
                }
            }

            public void SetPinLastBusPoint(int pin)
            {
                if (BusPoints != null)
                {
                    BusPoint bP = BusPoints[BusPoints.Count - 1];
                    bP.Pin = pin;
                    BusPoints[BusPoints.Count - 1] = bP;
                }
            }
        }

        struct BusPoint
        {
            Point3d point;
            bool end;
            bool page;
            int pin;

            public Point3d Point
            {
                get { return point; }
                set { point = value; }
            }
            public bool End
            {
                get { return end; }
                set { end = value; }
            }
            public bool Page
            {
                get { return page; }
                set { page = value; }
            }
            public int Pin
            {
                get { return pin; }
                set { pin = value; }
            }
        }

        struct UnitD
        {
            public string XT;
            public string XR;
            public string Equipment;
            public string Article;
            public int Rack;
            public int Slot;
            public int Channel;
            public string Type;
            public List<string> XRs;
            public int XRNum;
            public string RightPin;
            public int LeftPin;
            public List<string> XTPins;

            public int LinesNum;
            public string XTLink;
            public bool Table;
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
        }

        struct CCUnit
        {
            public string designation;
            public List<string> terminalsIn;
            public List<string> terminalsOut;
            public string param;
        }

        struct Group
        {
            public List<UnitA> UnitsA;
            public List<UnitD> UnitsD;
            public Unit? unit;
            public CCUnit? ccunit;
            public List<Point3d> Lines;
        }

        public static void DrawScheme()
        {
            UnitsA = new List<UnitA>();
            UnitsD = new List<UnitD>();
            Units = new List<Unit?>();
            CCUnits = new List<CCUnit?>();
            busXR = new List<Bus>();
            busXT = new List<Bus>();

            List<string> data = new List<string>();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "(*.txt)|*.txt";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                data = System.IO.File.ReadAllLines(ofd.FileName).ToList();
                for (int i = 0; i < data.Count; i++)
                {
                    string[] str = data[i].Split('░');
                    if (str[0] == "AI" || str[0] == "AO")
                    {
                        UnitA unitA = new UnitA();
                        unitA.Type = str[0];
                        unitA.Article = str[1];
                        unitA.Equipment = str[2];
                        unitA.Rack = int.Parse(str[3]);
                        unitA.Slot = int.Parse(str[4]);
                        unitA.Channel = int.Parse(str[5]);
                        unitA.Terminals = str[6];
                        string[] fu = str[7].Replace("{", string.Empty).Replace("}", string.Empty).Trim().Split(' ');
                        if (fu.Length > 0)
                            if (fu[0] != "")
                                unitA.Fu = new List<string>(fu);
                        string[] pins = str[8].Replace("{", string.Empty).Replace("}", string.Empty).Trim().Split(' ');
                        if (pins.Length > 0)
                            if (pins[0] != "")
                                unitA.Pins = new List<string>(pins);
                        UnitsA.Add(unitA);
                    }
                    else
                        if (str[0] == "DO" && str[6] == "-XT24.7")
                        {
                            UnitD unitD = new UnitD();
                            unitD.Type = str[0];
                            unitD.Article = str[1];
                            unitD.Equipment = str[2];
                            unitD.Rack = int.Parse(str[3]);
                            unitD.Slot = int.Parse(str[4]);
                            unitD.Channel = int.Parse(str[5]);
                            unitD.XT = str[6];
                            string[] xtpins = str[7].Replace("{", string.Empty).Replace("}", string.Empty).Trim().Split(' ');
                            if (xtpins.Length > 0)
                                if (xtpins[0] != "")
                                    unitD.XTPins = new List<string>(xtpins);
                            unitD.Table = false;
                            UnitsD.Add(unitD);
                        }
                        else
                        {
                            UnitD unitD = new UnitD();
                            unitD.Type = str[0];
                            unitD.Article = str[1];
                            unitD.Equipment = str[2];
                            unitD.Rack = int.Parse(str[3]);
                            unitD.Slot = int.Parse(str[4]);
                            unitD.Channel = int.Parse(str[5]);
                            unitD.RightPin = str[6];
                            unitD.LeftPin = int.Parse(str[7]);
                            unitD.XR = str[8];
                            unitD.XRNum = int.Parse(str[9]);
                            string[] fcs = str[10].Replace("{", string.Empty).Replace("}", string.Empty).Trim().Split(' ');
                            if (fcs.Length > 0)
                                if (fcs[0] != "")
                                    unitD.XRs = new List<string>(fcs);
                            unitD.XT = str[11];
                            string[] xtpins = str[12].Replace("{", string.Empty).Replace("}", string.Empty).Trim().Split(' ');
                            if (xtpins.Length > 0)
                                if (xtpins[0] != "")
                                    unitD.XTPins = new List<string>(xtpins);
                            unitD.XTLink = str[13];
                            unitD.Table = true;
                            UnitsD.Add(unitD);
                        }
                }
                OpenFileDialog ofd2 = new OpenFileDialog();
                ofd2.Filter = "(*.xlsx; *.xls)|*.xlsx; *.xls";
                if (ofd2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Units = loadDataFromConnectionScheme(ofd2.FileName);
                    OpenFileDialog ofd3 = new OpenFileDialog();
                    ofd3.Filter = "(*.xlsx; *.xls)|*.xlsx; *.xls";
                    if (ofd3.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        CCUnits = loadDataFromCupboardConnections(ofd3.FileName);
                        Groups = FindGroups();

                        acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                        Database acDb = acDoc.Database;
                        using (DocumentLock docLock = acDoc.LockDocument())
                        {
                            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
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
                                tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForRead);
                                loadFonts(acTrans, acModSpace, acDb);

                                PromptIntegerResult cNum = editor.GetInteger("\nВведите номер шкафа: ");
                                if (cNum.Status != PromptStatus.OK)
                                {
                                    editor.WriteMessage("Неверный ввод...");
                                    return;
                                }
                                else CupboardNum = cNum.Value;

                                PromptPointResult selectedPoint = acDoc.Editor.GetPoint("Выберите точку");
                                if (selectedPoint.Status == PromptStatus.OK)
                                {
                                    int pageNum = 1;
                                    Point3d point = DrawFrame(acTrans, acModSpace, selectedPoint.Value, pageNum);
                                    pageNum++;
                                    Point3d startPoint = point;
                                    point = point.Add(new Vector3d(-10, -35, 0));
                                    string prevTerm = "";
                                    bool terminal = false;
                                    bool page = true;
                                    bool union = false;
                                    Polyline termPoly = null;
                                    int unitCount = 0;
                                    for (int i = 0; i < Groups.Count; i++)
                                    {
                                        int count = 0;
                                        if (Groups[i].UnitsA != null) count += Groups[i].UnitsA.Count;
                                        if (Groups[i].UnitsD != null) count += Groups[i].UnitsD.Count;
                                        if (point.Y - (count) * 55 < startPoint.Y - 200)
                                        {
                                            point = DrawFrame(acTrans, acModSpace, startPoint, pageNum);
                                            pageNum++;
                                            startPoint = point;
                                            point = point.Add(new Vector3d(-10, -35, 0));
                                            page = true;
                                            union = false;
                                            unitCount = 0;
                                        }
                                        else
                                        {
                                            page = false;
                                            point = point.Add(new Vector3d(0, -25, 0));
                                        }
                                        if (Groups[i].UnitsA != null)
                                        {
                                            for (int j = 0; j < Groups[i].UnitsA.Count; j++)
                                            {
                                                if (prevTerm == "" || prevTerm != Groups[i].UnitsA[j].Terminals || page)
                                                {
                                                    terminal = true;
                                                    union = false;
                                                    unitCount = 1;
                                                }
                                                else
                                                {
                                                    terminal = false;
                                                    if (unitCount == 1)
                                                        union = true;
                                                    unitCount++;
                                                }
                                                double height = 0;
                                                termPoly = DrawA(acTrans, acModSpace, point, i, j, terminal, out height, termPoly, union);
                                                union = false;
                                                prevTerm = Groups[i].UnitsA[j].Terminals;
                                                point = point.Add(new Vector3d(0, -height - 15, 0));
                                            }
                                        }
                                        if (Groups[i].UnitsD != null)
                                        {
                                            if (Groups[i].UnitsD.Count > 1)
                                            { }
                                            for (int j = 0; j < Groups[i].UnitsD.Count; j++)
                                            {
                                                if ((prevTerm == "" || prevTerm != Groups[i].UnitsD[j].XT || page))
                                                {
                                                    terminal = true;

                                                    if (Groups[i].UnitsD[j].Table)
                                                    {
                                                        busXR.Add(new Bus());
                                                        Bus bXR = busXR[busXR.Count - 1];
                                                        bXR.Term = Groups[i].UnitsD[j].XR;
                                                        busXR[busXR.Count - 1] = bXR;

                                                        busXT.Add(new Bus());
                                                        Bus bXT = busXT[busXT.Count - 1];
                                                        bXT.Term = Groups[i].UnitsD[j].XT;
                                                        bXT.Link = Groups[i].UnitsD[j].XTLink;
                                                        bXT.Type = Groups[i].UnitsD[j].Type;
                                                        busXT[busXT.Count - 1] = bXT;
                                                    }
                                                    page = false;
                                                    if (j > 0) point = point.Add(new Vector3d(0, -35, 0));
                                                }
                                                else terminal = false;
                                                DrawD(acTrans, acModSpace, point, i, j, terminal);
                                                prevTerm = Groups[i].UnitsD[j].XT;
                                                point = point.Add(new Vector3d(0, -35, 0));
                                            }
                                        }
                                        point = point.Add(new Vector3d(0, -20, 0));
                                    }
                                    DrawBus(acTrans, acModSpace);
                                    DrawConnectionSchemePart(acTrans, acModSpace);
                                    acTrans.Commit();
                                }
                            }
                        }
                    }
                }
            }
        }

        private static Point3d DrawFrame(Transaction acTrans, BlockTableRecord modSpace, Point3d point, int pageNum)
        {
            Polyline framePoly = new Polyline();
            framePoly.SetDatabaseDefaults();
            framePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 3);
            framePoly.Closed = true;
            framePoly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
            framePoly.AddVertexAt(1, new Point2d(point.X+420, point.Y), 0, 0, 0);
            framePoly.AddVertexAt(2, new Point2d(point.X+420, point.Y-292), 0, 0, 0);
            framePoly.AddVertexAt(3, new Point2d(point.X, point.Y-292), 0, 0, 0);
            modSpace.AppendEntity(framePoly);
            acTrans.AddNewlyCreatedDBObject(framePoly, true);

            DBText pageNumTxt = new DBText();
            pageNumTxt.SetDatabaseDefaults();
            if (tst.Has("ROMANS0-60"))
            {
                pageNumTxt.TextStyleId = tst["ROMANS0-60"];
            }
            pageNumTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            pageNumTxt.Justify = AttachmentPoint.BottomMid;
            pageNumTxt.TextString = pageNum.ToString();
            pageNumTxt.Height = 50;
            pageNumTxt.Position = framePoly.GetPoint3dAt(1).Add(new Vector3d(-40, -260, 0));
            pageNumTxt.HorizontalMode = TextHorizontalMode.TextMid;
            pageNumTxt.AlignmentPoint = pageNumTxt.Position;
            modSpace.AppendEntity(pageNumTxt);
            acTrans.AddNewlyCreatedDBObject(pageNumTxt, true);

            return framePoly.GetPoint3dAt(1);
        }

        private static Polyline DrawA(Transaction acTrans, BlockTableRecord modSpace, Point3d point, int groupIndex, int unitIndex, bool terminal, out double Height, Polyline unionPoly, bool union)
        {
            double height = 0;

            UnitA unitA = Groups[groupIndex].UnitsA[unitIndex];
            unitA.PinPoint = new List<Point3d>();

            Group group = Groups[groupIndex];
            if (group.Lines == null) group.Lines = new List<Point3d>();
       
            DrawTableRow(acTrans, modSpace, point, "НАЗВ. ПЛК / PLC NAME : PCS 7");
            point = point.Add(new Vector3d(0, -4.12, 0));
            height += 4.12;
            DrawTableRow(acTrans, modSpace, point, "МОДУЛЬ / I/O MODULE : " + unitA.Article);
            point = point.Add(new Vector3d(0, -4.12, 0));
            height += 4.12;
            DrawTableRow(acTrans, modSpace, point, "КОРЗИНА / RACK : " + unitA.Rack);
            point = point.Add(new Vector3d(0, -4.12, 0));
            height += 4.12;
            DrawTableRow(acTrans, modSpace, point, "ЯЧЕЙКА / SLOT : " + unitA.Slot);
            point = point.Add(new Vector3d(0, -4.12, 0));
            height += 4.12;
            DrawTableRow(acTrans, modSpace, point, "КАНАЛ / CHANNEL : " + unitA.Channel);
            point = point.Add(new Vector3d(0, -4.12, 0));
            height += 4.12;

            Polyline plcTypePoly = new Polyline();
            plcTypePoly.SetDatabaseDefaults();
            plcTypePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
            plcTypePoly.Closed = true;
            plcTypePoly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
            plcTypePoly.AddVertexAt(1, new Point2d(point.X, point.Y - 5 * unitA.Pins.Count), 0, 0, 0);
            plcTypePoly.AddVertexAt(2, new Point2d(point.X - 40, point.Y - 5 * unitA.Pins.Count), 0, 0, 0);
            plcTypePoly.AddVertexAt(3, new Point2d(point.X - 40, point.Y), 0, 0, 0);
            modSpace.AppendEntity(plcTypePoly);
            acTrans.AddNewlyCreatedDBObject(plcTypePoly, true);
            height += 5 * unitA.Pins.Count;

            DBText plcTypeTxt = new DBText();
            plcTypeTxt.SetDatabaseDefaults();
            plcTypeTxt.WidthFactor = 0.5;
            if (tst.Has("ROMANS0-60"))
            {
                plcTypeTxt.TextStyleId = tst["ROMANS0-60"];
            }
            plcTypeTxt.Color = Color.FromRgb(165, 82, 0);
            plcTypeTxt.Justify = AttachmentPoint.MiddleLeft;
            plcTypeTxt.TextString = "ТИП / TYPE : " + unitA.Type;
            plcTypeTxt.Height = 2.5;
            plcTypeTxt.Position = plcTypePoly.GetPoint3dAt(3).Add(new Vector3d(1, (-5 * unitA.Pins.Count) / 2, 0));
            plcTypeTxt.HorizontalMode = TextHorizontalMode.TextLeft;
            plcTypeTxt.AlignmentPoint = plcTypeTxt.Position;
            modSpace.AppendEntity(plcTypeTxt);
            acTrans.AddNewlyCreatedDBObject(plcTypeTxt, true);

            point = point.Add(new Vector3d(-40, 0, 0));
            List<string> pinsAI = new List<string>();
            pinsAI.Add("M" + unitA.Channel + "+");
            pinsAI.Add("M" + unitA.Channel + "-");

            List<string> pinsAO = new List<string>();
            pinsAO.Add("QV" + unitA.Channel + "+");
            pinsAO.Add("S" + unitA.Channel + "+");
            pinsAO.Add("S" + unitA.Channel + "-");
            pinsAO.Add("M ana");

            List<Point3d> points = new List<Point3d>();
            List<Point3d> linePoint = new List<Point3d>();
            for (int i = 0; i < unitA.Pins.Count; i++)
            {
                Polyline pinPoly = new Polyline();
                pinPoly.SetDatabaseDefaults();
                pinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                pinPoly.Closed = true;
                pinPoly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
                pinPoly.AddVertexAt(1, new Point2d(point.X, point.Y - 5), 0, 0, 0);
                pinPoly.AddVertexAt(2, new Point2d(point.X - 11.5, point.Y - 5), 0, 0, 0);
                pinPoly.AddVertexAt(3, new Point2d(point.X - 11.5, point.Y), 0, 0, 0);
                modSpace.AppendEntity(pinPoly);
                acTrans.AddNewlyCreatedDBObject(pinPoly, true);

                DBText pinNameTxt = new DBText();
                pinNameTxt.SetDatabaseDefaults();
                pinNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    pinNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                pinNameTxt.Color = Color.FromRgb(165, 82, 0);
                pinNameTxt.Justify = AttachmentPoint.MiddleLeft;
                pinNameTxt.TextString = unitA.Type == "AI" ? pinsAI[i] : pinsAO[i];
                pinNameTxt.Height = 2.5;
                pinNameTxt.Position = pinPoly.GetPoint3dAt(3).Add(new Vector3d(1, -2.06, 0));
                pinNameTxt.HorizontalMode = TextHorizontalMode.TextLeft;
                pinNameTxt.AlignmentPoint = pinNameTxt.Position;
                modSpace.AppendEntity(pinNameTxt);
                acTrans.AddNewlyCreatedDBObject(pinNameTxt, true);

                Polyline leftPartPinPoly = new Polyline();
                leftPartPinPoly.SetDatabaseDefaults();
                leftPartPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                leftPartPinPoly.Closed = true;
                leftPartPinPoly.AddVertexAt(0, pinPoly.GetPoint2dAt(3), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(1, leftPartPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(2, leftPartPinPoly.GetPoint2dAt(1).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(3, leftPartPinPoly.GetPoint2dAt(2).Add(new Vector2d(0, 5)), 0, 0, 0);
                modSpace.AppendEntity(leftPartPinPoly);
                acTrans.AddNewlyCreatedDBObject(leftPartPinPoly, true);

                Line lineUp = new Line();
                lineUp.SetDatabaseDefaults();
                lineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineUp.StartPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(6, 0, 0));
                lineUp.EndPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(7, -2.5, 0));
                modSpace.AppendEntity(lineUp);
                acTrans.AddNewlyCreatedDBObject(lineUp, true);

                Line lineDown = new Line();
                lineDown.SetDatabaseDefaults();
                lineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineDown.StartPoint = leftPartPinPoly.GetPoint3dAt(2).Add(new Vector3d(7, 2.5, 0));
                lineDown.EndPoint = leftPartPinPoly.GetPoint3dAt(2).Add(new Vector3d(6, 0, 0));
                modSpace.AppendEntity(lineDown);
                acTrans.AddNewlyCreatedDBObject(lineDown, true);

                Line lineFromTable = new Line();
                lineFromTable.SetDatabaseDefaults();
                lineFromTable.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineFromTable.StartPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0));
                lineFromTable.EndPoint = lineFromTable.StartPoint.Add(new Vector3d(-6, 0, 0));
                modSpace.AppendEntity(lineFromTable);
                acTrans.AddNewlyCreatedDBObject(lineFromTable, true);

                Polyline rightPartPinPoly = new Polyline();
                rightPartPinPoly.SetDatabaseDefaults();
                rightPartPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                rightPartPinPoly.Closed = true;
                rightPartPinPoly.AddVertexAt(0, new Point2d(lineFromTable.EndPoint.X, lineFromTable.EndPoint.Y+2.5), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(1, rightPartPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(2, rightPartPinPoly.GetPoint2dAt(1).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(3, rightPartPinPoly.GetPoint2dAt(2).Add(new Vector2d(0, 5)), 0, 0, 0);
                modSpace.AppendEntity(rightPartPinPoly);
                acTrans.AddNewlyCreatedDBObject(rightPartPinPoly, true);

                Line lineUp2 = new Line();
                lineUp2.SetDatabaseDefaults();
                lineUp2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineUp2.StartPoint = rightPartPinPoly.GetPoint3dAt(0).Add(new Vector3d(-6, 0, 0));
                lineUp2.EndPoint = rightPartPinPoly.GetPoint3dAt(0).Add(new Vector3d(-7, -2.5, 0));
                modSpace.AppendEntity(lineUp2);
                acTrans.AddNewlyCreatedDBObject(lineUp2, true);
                points.Add(lineUp2.StartPoint);

                Line lineDown2 = new Line();
                lineDown2.SetDatabaseDefaults();
                lineDown2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineDown2.StartPoint = rightPartPinPoly.GetPoint3dAt(1).Add(new Vector3d(-7, 2.5, 0));
                lineDown2.EndPoint = rightPartPinPoly.GetPoint3dAt(1).Add(new Vector3d(-6, 0, 0));
                modSpace.AppendEntity(lineDown2);
                acTrans.AddNewlyCreatedDBObject(lineDown2, true);

                Polyline leftPinPoly = new Polyline();
                leftPinPoly.SetDatabaseDefaults();
                leftPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                leftPinPoly.Closed = true;
                leftPinPoly.AddVertexAt(0, rightPartPinPoly.GetPoint2dAt(3).Add(new Vector2d(-5, 0)), 0, 0, 0);
                leftPinPoly.AddVertexAt(1, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                leftPinPoly.AddVertexAt(2, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(-11.5, -5)), 0, 0, 0);
                leftPinPoly.AddVertexAt(3, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                modSpace.AppendEntity(leftPinPoly);
                acTrans.AddNewlyCreatedDBObject(leftPinPoly, true);
                points.Add(new Point3d(leftPinPoly.GetPoint3dAt(3).X + (leftPinPoly.GetPoint3dAt(0).X - leftPinPoly.GetPoint3dAt(3).X) / 2, leftPinPoly.GetPoint3dAt(0).Y + 2.5, 0));

                DBText leftPinNameTxt = new DBText();
                leftPinNameTxt.SetDatabaseDefaults();
                leftPinNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    leftPinNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                leftPinNameTxt.Color = Color.FromRgb(165, 82, 0);
                leftPinNameTxt.Justify = AttachmentPoint.MiddleLeft;
                leftPinNameTxt.TextString = unitA.Pins[i];
                leftPinNameTxt.Height = 2.5;
                leftPinNameTxt.Position = leftPinPoly.GetPoint3dAt(3).Add(new Vector3d(5, -2.06, 0));
                leftPinNameTxt.HorizontalMode = TextHorizontalMode.TextLeft;
                leftPinNameTxt.AlignmentPoint = leftPinNameTxt.Position;
                modSpace.AppendEntity(leftPinNameTxt);
                acTrans.AddNewlyCreatedDBObject(leftPinNameTxt, true);

                linePoint.Add(leftPinPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0)));

                point = point.Add(new Vector3d(0, -5, 0));
            }

            if (linePoint.Count==2)
            {
                for (int i=0; i<linePoint.Count; i++)
                {
                    Line lineFromPin = new Line();
                    lineFromPin.SetDatabaseDefaults();
                    lineFromPin.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                    lineFromPin.StartPoint = linePoint[i];
                    lineFromPin.EndPoint = lineFromPin.StartPoint.Add(new Vector3d(-25, 0, 0));
                    modSpace.AppendEntity(lineFromPin);
                    acTrans.AddNewlyCreatedDBObject(lineFromPin, true);

                    unitA.PinPoint.Add(lineFromPin.EndPoint);
                    Groups[groupIndex].UnitsA[unitIndex] = unitA;
                }
            }
            else
            {
                Line lineFromPin1 = new Line();
                lineFromPin1.SetDatabaseDefaults();
                lineFromPin1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromPin1.StartPoint = linePoint[0];
                lineFromPin1.EndPoint = lineFromPin1.StartPoint.Add(new Vector3d(-25, 0, 0));
                modSpace.AppendEntity(lineFromPin1);
                acTrans.AddNewlyCreatedDBObject(lineFromPin1, true);

                unitA.PinPoint.Add(lineFromPin1.EndPoint);

                Polyline lineFromPin2 = new Polyline();
                lineFromPin2.SetDatabaseDefaults();
                lineFromPin2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromPin2.Closed = false;
                lineFromPin2.AddVertexAt(0, new Point2d(linePoint[0].X, linePoint[0].Y), 0, 0, 0);
                lineFromPin2.AddVertexAt(1, lineFromPin2.GetPoint2dAt(0).Add(new Vector2d(-2.8, -1.8)), 0, 0, 0);
                lineFromPin2.AddVertexAt(2, lineFromPin2.GetPoint2dAt(1).Add(new Vector2d(0, -1.2)), 0, 0, 0);
                lineFromPin2.AddVertexAt(3, new Point2d(linePoint[1].X, linePoint[1].Y), 0, 0, 0);
                modSpace.AppendEntity(lineFromPin2);
                acTrans.AddNewlyCreatedDBObject(lineFromPin2, true);

                Polyline lineFromPin3 = new Polyline();
                lineFromPin3.SetDatabaseDefaults();
                lineFromPin3.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromPin3.Closed = false;
                lineFromPin3.AddVertexAt(0, new Point2d(linePoint[2].X, linePoint[2].Y), 0, 0, 0);
                lineFromPin3.AddVertexAt(1, lineFromPin3.GetPoint2dAt(0).Add(new Vector2d(-2.8, -1.8)), 0, 0, 0);
                lineFromPin3.AddVertexAt(2, lineFromPin3.GetPoint2dAt(1).Add(new Vector2d(0, -1.2)), 0, 0, 0);
                lineFromPin3.AddVertexAt(3, new Point2d(linePoint[3].X, linePoint[3].Y), 0, 0, 0);
                modSpace.AppendEntity(lineFromPin3);
                acTrans.AddNewlyCreatedDBObject(lineFromPin3, true);

                Polyline lineFromPin4 = new Polyline();
                lineFromPin4.SetDatabaseDefaults();
                lineFromPin4.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromPin4.Closed = false;
                lineFromPin4.AddVertexAt(0, new Point2d(linePoint[3].X, linePoint[3].Y), 0, 0, 0);
                lineFromPin4.AddVertexAt(1, lineFromPin4.GetPoint2dAt(0).Add(new Vector2d(-12.5, 0)), 0, 0, 0);
                lineFromPin4.AddVertexAt(2, lineFromPin4.GetPoint2dAt(1).Add(new Vector2d(0, 10)), 0, 0, 0);
                lineFromPin4.AddVertexAt(3, lineFromPin4.GetPoint2dAt(2).Add(new Vector2d(-12.5, 0)), 0, 0, 0);
                modSpace.AppendEntity(lineFromPin4);
                acTrans.AddNewlyCreatedDBObject(lineFromPin4, true);

                unitA.PinPoint.Add(lineFromPin4.EndPoint);
                group.UnitsA[unitIndex] = unitA;
            }

            Point2d termPoint = Point2d.Origin;
            Point2d lastTermPoint = Point2d.Origin;
            for (int i = 0; i < unitA.Fu.Count; i++)
            {
                Polyline fuPoly = new Polyline();
                fuPoly.SetDatabaseDefaults();
                fuPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                fuPoly.Closed = true;
                fuPoly.AddVertexAt(0, new Point2d(unitA.PinPoint[i].X, unitA.PinPoint[i].Y + 2.5), 0, 0, 0);
                fuPoly.AddVertexAt(1, fuPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                fuPoly.AddVertexAt(2, fuPoly.GetPoint2dAt(0).Add(new Vector2d(-13, -5)), 0, 0, 0);
                fuPoly.AddVertexAt(3, fuPoly.GetPoint2dAt(0).Add(new Vector2d(-13, 0)), 0, 0, 0);
                modSpace.AppendEntity(fuPoly);
                acTrans.AddNewlyCreatedDBObject(fuPoly, true);

                DBText fuNameTxt = new DBText();
                fuNameTxt.SetDatabaseDefaults();
                fuNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    fuNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                fuNameTxt.Color = Color.FromRgb(165, 82, 0);
                fuNameTxt.Justify = AttachmentPoint.MiddleCenter;
                fuNameTxt.TextString = unitA.Fu[i];
                fuNameTxt.Height = 2.5;
                fuNameTxt.Position = fuPoly.GetPoint3dAt(3).Add(new Vector3d(6.5, -2.06, 0));
                fuNameTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                fuNameTxt.AlignmentPoint = fuNameTxt.Position;
                modSpace.AppendEntity(fuNameTxt);
                acTrans.AddNewlyCreatedDBObject(fuNameTxt, true);

                Line lineFromFu = new Line();
                lineFromFu.SetDatabaseDefaults();
                lineFromFu.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                lineFromFu.StartPoint = fuPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0));
                lineFromFu.EndPoint = lineFromFu.StartPoint.Add(new Vector3d(-30, 0, 0));
                modSpace.AppendEntity(lineFromFu);
                acTrans.AddNewlyCreatedDBObject(lineFromFu, true);

                group.Lines.Add(lineFromFu.EndPoint);

                if (i==0) termPoint = fuPoly.GetPoint2dAt(3);
                else lastTermPoint = fuPoly.GetPoint2dAt(2);
            }

            Polyline termUnionPoly = null;
            if (terminal)
            {
                Polyline termPoly = new Polyline();
                termPoly.SetDatabaseDefaults();
                termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                termPoly.Closed = true;
                termPoly.AddVertexAt(0, termPoint.Add(new Vector2d(-3, 10)), 0, 0, 0);
                termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(0, 10)), 0, 0, 0);
                termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(0).Add(new Vector2d(20, 10)), 0, 0, 0);
                termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(0).Add(new Vector2d(20, 0)), 0, 0, 0);
                modSpace.AppendEntity(termPoly);
                acTrans.AddNewlyCreatedDBObject(termPoly, true);

                DBText termNameTxt = new DBText();
                termNameTxt.SetDatabaseDefaults();
                termNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    termNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                termNameTxt.Color = Color.FromRgb(165, 82, 0);
                termNameTxt.Justify = AttachmentPoint.MiddleCenter;
                termNameTxt.TextString = unitA.Terminals;
                termNameTxt.Height = 2.5;
                termNameTxt.Position = termPoly.GetPoint3dAt(1).Add(new Vector3d(10, -5, 0));
                termNameTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                termNameTxt.AlignmentPoint = termNameTxt.Position;
                modSpace.AppendEntity(termNameTxt);
                acTrans.AddNewlyCreatedDBObject(termNameTxt, true);

                termUnionPoly = new Polyline();
                termUnionPoly.SetDatabaseDefaults();
                termUnionPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                termUnionPoly.Closed = true;
                termUnionPoly.AddVertexAt(0, termPoint, 0, 0, 0);
                termUnionPoly.AddVertexAt(1, termUnionPoly.GetPoint2dAt(0).Add(new Vector2d(13, 0)), 0, 0, 0);
                termUnionPoly.AddVertexAt(2, termUnionPoly.GetPoint2dAt(0).Add(new Vector2d(13, -5)), 0, 0, 0);
                termUnionPoly.AddVertexAt(3, termUnionPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                modSpace.AppendEntity(termUnionPoly);
                acTrans.AddNewlyCreatedDBObject(termUnionPoly, true);
            }
            else
            {
                termUnionPoly = unionPoly;
                if (union)
                {
                    termUnionPoly.SetPointAt(0, termUnionPoly.GetPoint2dAt(0).Add(new Vector2d(0, 4)));
                    termUnionPoly.SetPointAt(1, termUnionPoly.GetPoint2dAt(1).Add(new Vector2d(0, 4)));
                    union = false;
                }
                termUnionPoly.SetPointAt(2, lastTermPoint.Add(new Vector2d(13, 0)));
                termUnionPoly.SetPointAt(3, lastTermPoint);
            }

            if (points.Count >= 2)
            {
                Polyline jumperUp = new Polyline();
                jumperUp.SetDatabaseDefaults();
                jumperUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                jumperUp.Closed = false;
                jumperUp.AddVertexAt(0, new Point2d(points[0].X, points[0].Y), 0, 0, 0);
                jumperUp.AddVertexAt(1, jumperUp.GetPoint2dAt(0).Add(new Vector2d(0, 2.5)), 0, 0, 0);
                jumperUp.AddVertexAt(2, new Point2d(points[1].X, points[1].Y), 0, 0, 0);
                jumperUp.AddVertexAt(3, jumperUp.GetPoint2dAt(2).Add(new Vector2d(0, -2.5)), 0, 0, 0);
                modSpace.AppendEntity(jumperUp);
                acTrans.AddNewlyCreatedDBObject(jumperUp, true);

                Polyline jumperDown = new Polyline();
                jumperDown.SetDatabaseDefaults();
                jumperDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                jumperDown.Closed = false;
                jumperDown.AddVertexAt(0, jumperUp.GetPoint2dAt(0).Add(new Vector2d(0, -5 * unitA.Pins.Count)), 0, 0, 0);
                jumperDown.AddVertexAt(1, jumperDown.GetPoint2dAt(0).Add(new Vector2d(0, -2.5)), 0, 0, 0);
                jumperDown.AddVertexAt(2, jumperUp.GetPoint2dAt(3).Add(new Vector2d(0, -5 * unitA.Pins.Count - 2.5)), 0, 0, 0);
                jumperDown.AddVertexAt(3, jumperDown.GetPoint2dAt(2).Add(new Vector2d(0, 2.5)), 0, 0, 0);
                modSpace.AppendEntity(jumperDown);
                acTrans.AddNewlyCreatedDBObject(jumperDown, true);
            }

            Groups[groupIndex] = group;
            Height = height;
            return termUnionPoly;
        }

        private static void DrawD(Transaction acTrans, BlockTableRecord modSpace, Point3d point, int groupIndex, int unitIndex, bool terminal)
        {
            double height = 0;
            UnitD unitD = Groups[groupIndex].UnitsD[unitIndex];
            Group group = Groups[groupIndex];
            if (group.Lines == null) group.Lines = new List<Point3d>();

            Point3d Point = Point3d.Origin;
            Point2d fcsPoint = Point2d.Origin;
            Point2d busPoint = new Point2d();
            List<Point3d> jumperPoints = new List<Point3d>();
            Line lineFromFcs = new Line();
            if (unitD.Table)
            {
                DrawTableRow(acTrans, modSpace, point, "НАЗВ. ПЛК / PLC NAME : PCS 7");
                point = point.Add(new Vector3d(0, -4.12, 0));
                height += 4.12;
                DrawTableRow(acTrans, modSpace, point, "МОДУЛЬ / I/O MODULE : " + unitD.Article);
                point = point.Add(new Vector3d(0, -4.12, 0));
                height += 4.12;
                DrawTableRow(acTrans, modSpace, point, "КОРЗИНА / RACK : " + unitD.Rack);
                point = point.Add(new Vector3d(0, -4.12, 0));
                height += 4.12;
                DrawTableRow(acTrans, modSpace, point, "ЯЧЕЙКА / SLOT : " + unitD.Slot);
                point = point.Add(new Vector3d(0, -4.12, 0));
                height += 4.12;
                DrawTableRow(acTrans, modSpace, point, "КАНАЛ / CHANNEL : " + unitD.Channel);
                point = point.Add(new Vector3d(0, -4.12, 0));
                height += 4.12;

                Polyline plcTypePoly = new Polyline();
                plcTypePoly.SetDatabaseDefaults();
                plcTypePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                plcTypePoly.Closed = true;
                plcTypePoly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
                plcTypePoly.AddVertexAt(1, new Point2d(point.X, point.Y - 5), 0, 0, 0);
                plcTypePoly.AddVertexAt(2, new Point2d(point.X - 40, point.Y - 5), 0, 0, 0);
                plcTypePoly.AddVertexAt(3, new Point2d(point.X - 40, point.Y), 0, 0, 0);
                modSpace.AppendEntity(plcTypePoly);
                acTrans.AddNewlyCreatedDBObject(plcTypePoly, true);
                height += 5;

                DBText plcTypeTxt = new DBText();
                plcTypeTxt.SetDatabaseDefaults();
                plcTypeTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    plcTypeTxt.TextStyleId = tst["ROMANS0-60"];
                }
                plcTypeTxt.Color = Color.FromRgb(165, 82, 0);
                plcTypeTxt.Justify = AttachmentPoint.MiddleLeft;
                plcTypeTxt.TextString = "ТИП / TYPE : " + unitD.Type;
                plcTypeTxt.Height = 2.5;
                plcTypeTxt.Position = plcTypePoly.GetPoint3dAt(3).Add(new Vector3d(1, -(5 / 2), 0));
                plcTypeTxt.HorizontalMode = TextHorizontalMode.TextLeft;
                plcTypeTxt.AlignmentPoint = plcTypeTxt.Position;
                modSpace.AppendEntity(plcTypeTxt);
                acTrans.AddNewlyCreatedDBObject(plcTypeTxt, true);

                point = point.Add(new Vector3d(-40, 0, 0));

                Polyline pinPoly = new Polyline();
                pinPoly.SetDatabaseDefaults();
                pinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                pinPoly.Closed = true;
                pinPoly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
                pinPoly.AddVertexAt(1, new Point2d(point.X, point.Y - 5), 0, 0, 0);
                pinPoly.AddVertexAt(2, new Point2d(point.X - 11.5, point.Y - 5), 0, 0, 0);
                pinPoly.AddVertexAt(3, new Point2d(point.X - 11.5, point.Y), 0, 0, 0);
                modSpace.AppendEntity(pinPoly);
                acTrans.AddNewlyCreatedDBObject(pinPoly, true);

                DBText pinNameTxt = new DBText();
                pinNameTxt.SetDatabaseDefaults();
                pinNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    pinNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                pinNameTxt.Color = Color.FromRgb(165, 82, 0);
                pinNameTxt.Justify = AttachmentPoint.MiddleLeft;
                pinNameTxt.TextString = unitD.RightPin;
                pinNameTxt.Height = 2.5;
                pinNameTxt.Position = pinPoly.GetPoint3dAt(3).Add(new Vector3d(1, -2.06, 0));
                pinNameTxt.HorizontalMode = TextHorizontalMode.TextLeft;
                pinNameTxt.AlignmentPoint = pinNameTxt.Position;
                modSpace.AppendEntity(pinNameTxt);
                acTrans.AddNewlyCreatedDBObject(pinNameTxt, true);

                Polyline leftPartPinPoly = new Polyline();
                leftPartPinPoly.SetDatabaseDefaults();
                leftPartPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                leftPartPinPoly.Closed = true;
                leftPartPinPoly.AddVertexAt(0, pinPoly.GetPoint2dAt(3), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(1, leftPartPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(2, leftPartPinPoly.GetPoint2dAt(1).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                leftPartPinPoly.AddVertexAt(3, leftPartPinPoly.GetPoint2dAt(2).Add(new Vector2d(0, 5)), 0, 0, 0);
                modSpace.AppendEntity(leftPartPinPoly);
                acTrans.AddNewlyCreatedDBObject(leftPartPinPoly, true);

                Line lineUp = new Line();
                lineUp.SetDatabaseDefaults();
                lineUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineUp.StartPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(6, 0, 0));
                lineUp.EndPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(7, -2.5, 0));
                modSpace.AppendEntity(lineUp);
                acTrans.AddNewlyCreatedDBObject(lineUp, true);

                Line lineDown = new Line();
                lineDown.SetDatabaseDefaults();
                lineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineDown.StartPoint = leftPartPinPoly.GetPoint3dAt(2).Add(new Vector3d(7, 2.5, 0));
                lineDown.EndPoint = leftPartPinPoly.GetPoint3dAt(2).Add(new Vector3d(6, 0, 0));
                modSpace.AppendEntity(lineDown);
                acTrans.AddNewlyCreatedDBObject(lineDown, true);

                Line lineFromTable = new Line();
                lineFromTable.SetDatabaseDefaults();
                lineFromTable.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineFromTable.StartPoint = leftPartPinPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0));
                lineFromTable.EndPoint = lineFromTable.StartPoint.Add(new Vector3d(-6, 0, 0));
                modSpace.AppendEntity(lineFromTable);
                acTrans.AddNewlyCreatedDBObject(lineFromTable, true);

                Polyline rightPartPinPoly = new Polyline();
                rightPartPinPoly.SetDatabaseDefaults();
                rightPartPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                rightPartPinPoly.Closed = true;
                rightPartPinPoly.AddVertexAt(0, new Point2d(lineFromTable.EndPoint.X, lineFromTable.EndPoint.Y + 2.5), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(1, rightPartPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(2, rightPartPinPoly.GetPoint2dAt(1).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                rightPartPinPoly.AddVertexAt(3, rightPartPinPoly.GetPoint2dAt(2).Add(new Vector2d(0, 5)), 0, 0, 0);
                modSpace.AppendEntity(rightPartPinPoly);
                acTrans.AddNewlyCreatedDBObject(rightPartPinPoly, true);

                Line lineUp2 = new Line();
                lineUp2.SetDatabaseDefaults();
                lineUp2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineUp2.StartPoint = rightPartPinPoly.GetPoint3dAt(0).Add(new Vector3d(-6, 0, 0));
                lineUp2.EndPoint = rightPartPinPoly.GetPoint3dAt(0).Add(new Vector3d(-7, -2.5, 0));
                modSpace.AppendEntity(lineUp2);
                acTrans.AddNewlyCreatedDBObject(lineUp2, true);

                Line lineDown2 = new Line();
                lineDown2.SetDatabaseDefaults();
                lineDown2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                lineDown2.StartPoint = rightPartPinPoly.GetPoint3dAt(1).Add(new Vector3d(-7, 2.5, 0));
                lineDown2.EndPoint = rightPartPinPoly.GetPoint3dAt(1).Add(new Vector3d(-6, 0, 0));
                modSpace.AppendEntity(lineDown2);
                acTrans.AddNewlyCreatedDBObject(lineDown2, true);

                Polyline leftPinPoly = new Polyline();
                leftPinPoly.SetDatabaseDefaults();
                leftPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                leftPinPoly.Closed = true;
                leftPinPoly.AddVertexAt(0, rightPartPinPoly.GetPoint2dAt(3).Add(new Vector2d(-5, 0)), 0, 0, 0);
                leftPinPoly.AddVertexAt(1, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                leftPinPoly.AddVertexAt(2, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(-11.5, -5)), 0, 0, 0);
                leftPinPoly.AddVertexAt(3, leftPinPoly.GetPoint2dAt(0).Add(new Vector2d(-11.5, 0)), 0, 0, 0);
                modSpace.AppendEntity(leftPinPoly);
                acTrans.AddNewlyCreatedDBObject(leftPinPoly, true);

                DBText leftPinNameTxt = new DBText();
                leftPinNameTxt.SetDatabaseDefaults();
                leftPinNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    leftPinNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                leftPinNameTxt.Color = Color.FromRgb(165, 82, 0);
                leftPinNameTxt.Justify = AttachmentPoint.MiddleLeft;
                leftPinNameTxt.TextString = unitD.LeftPin.ToString();
                leftPinNameTxt.Height = 2.5;
                leftPinNameTxt.Position = leftPinPoly.GetPoint3dAt(3).Add(new Vector3d(5, -2.06, 0));
                leftPinNameTxt.HorizontalMode = TextHorizontalMode.TextLeft;
                leftPinNameTxt.AlignmentPoint = leftPinNameTxt.Position;
                modSpace.AppendEntity(leftPinNameTxt);
                acTrans.AddNewlyCreatedDBObject(leftPinNameTxt, true);

                Polyline jumperUp = new Polyline();
                jumperUp.SetDatabaseDefaults();
                jumperUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                jumperUp.Closed = false;
                jumperUp.AddVertexAt(0, new Point2d(lineUp2.StartPoint.X, lineUp2.StartPoint.Y), 0, 0, 0);
                jumperUp.AddVertexAt(1, jumperUp.GetPoint2dAt(0).Add(new Vector2d(0, 2.5)), 0, 0, 0);
                Point3d p = new Point3d(leftPinPoly.GetPoint3dAt(3).X + (leftPinPoly.GetPoint3dAt(0).X - leftPinPoly.GetPoint3dAt(3).X) / 2, leftPinPoly.GetPoint3dAt(0).Y + 2.5, 0);
                jumperUp.AddVertexAt(2, new Point2d(p.X, p.Y), 0, 0, 0);
                jumperUp.AddVertexAt(3, jumperUp.GetPoint2dAt(2).Add(new Vector2d(0, -2.5)), 0, 0, 0);
                modSpace.AppendEntity(jumperUp);
                acTrans.AddNewlyCreatedDBObject(jumperUp, true);

                Polyline jumperDown = new Polyline();
                jumperDown.SetDatabaseDefaults();
                jumperDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                jumperDown.Closed = false;
                jumperDown.AddVertexAt(0, jumperUp.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                jumperDown.AddVertexAt(1, jumperDown.GetPoint2dAt(0).Add(new Vector2d(0, -2.5)), 0, 0, 0);
                jumperDown.AddVertexAt(2, jumperUp.GetPoint2dAt(3).Add(new Vector2d(0, -7.5)), 0, 0, 0);
                jumperDown.AddVertexAt(3, jumperDown.GetPoint2dAt(2).Add(new Vector2d(0, 2.5)), 0, 0, 0);
                modSpace.AppendEntity(jumperDown);
                acTrans.AddNewlyCreatedDBObject(jumperDown, true);

                point = point.Add(new Vector3d(0, -5, 0));

                Line lineFromPin = new Line();
                lineFromPin.SetDatabaseDefaults();
                lineFromPin.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromPin.StartPoint = leftPinPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0));
                lineFromPin.EndPoint = lineFromPin.StartPoint.Add(new Vector3d(-6, 0, 0));
                modSpace.AppendEntity(lineFromPin);
                acTrans.AddNewlyCreatedDBObject(lineFromPin, true);

                
                int FcsCount = 0;
                for (int i = 0; i < unitD.XRs.Count; i += 2)
                {
                    Polyline fcsPoly1 = new Polyline();
                    fcsPoly1.SetDatabaseDefaults();
                    fcsPoly1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    fcsPoly1.Closed = true;
                    fcsPoly1.AddVertexAt(0, new Point2d(lineFromPin.EndPoint.X, lineFromPin.EndPoint.Y + 2.5 - 5 * FcsCount), 0, 0, 0);
                    fcsPoly1.AddVertexAt(1, fcsPoly1.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                    fcsPoly1.AddVertexAt(2, fcsPoly1.GetPoint2dAt(0).Add(new Vector2d(-10, -5)), 0, 0, 0);
                    fcsPoly1.AddVertexAt(3, fcsPoly1.GetPoint2dAt(0).Add(new Vector2d(-10, 0)), 0, 0, 0);
                    modSpace.AppendEntity(fcsPoly1);
                    acTrans.AddNewlyCreatedDBObject(fcsPoly1, true);
                    jumperPoints.Add(new Point3d(fcsPoly1.GetPoint3dAt(3).X + (fcsPoly1.GetPoint3dAt(0).X - fcsPoly1.GetPoint3dAt(3).X) / 2, fcsPoly1.GetPoint3dAt(3).Y, 0));

                    DBText fcsNameTxt1 = new DBText();
                    fcsNameTxt1.SetDatabaseDefaults();
                    fcsNameTxt1.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        fcsNameTxt1.TextStyleId = tst["ROMANS0-60"];
                    }
                    fcsNameTxt1.Color = Color.FromRgb(165, 82, 0);
                    fcsNameTxt1.Justify = AttachmentPoint.MiddleCenter;
                    fcsNameTxt1.TextString = unitD.XRs[i + 1];
                    fcsNameTxt1.Height = 2.5;
                    fcsNameTxt1.Position = fcsPoly1.GetPoint3dAt(3).Add(new Vector3d(5, -2.06, 0));
                    fcsNameTxt1.HorizontalMode = TextHorizontalMode.TextCenter;
                    fcsNameTxt1.AlignmentPoint = fcsNameTxt1.Position;
                    modSpace.AppendEntity(fcsNameTxt1);
                    acTrans.AddNewlyCreatedDBObject(fcsNameTxt1, true);

                    Polyline fcsPoly2 = new Polyline();
                    fcsPoly2.SetDatabaseDefaults();
                    fcsPoly2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    fcsPoly2.Closed = true;
                    fcsPoly2.AddVertexAt(0, fcsPoly1.GetPoint2dAt(3).Add(new Vector2d(-5, 0)), 0, 0, 0);
                    fcsPoly2.AddVertexAt(1, fcsPoly2.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                    fcsPoly2.AddVertexAt(2, fcsPoly2.GetPoint2dAt(0).Add(new Vector2d(-10, -5)), 0, 0, 0);
                    fcsPoly2.AddVertexAt(3, fcsPoly2.GetPoint2dAt(0).Add(new Vector2d(-10, 0)), 0, 0, 0);
                    modSpace.AppendEntity(fcsPoly2);
                    acTrans.AddNewlyCreatedDBObject(fcsPoly2, true);
                    jumperPoints.Add(new Point3d(fcsPoly2.GetPoint3dAt(3).X + (fcsPoly2.GetPoint3dAt(0).X - fcsPoly2.GetPoint3dAt(3).X) / 2, fcsPoly2.GetPoint3dAt(3).Y, 0));

                    DBText fcsNameTxt2 = new DBText();
                    fcsNameTxt2.SetDatabaseDefaults();
                    fcsNameTxt2.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        fcsNameTxt2.TextStyleId = tst["ROMANS0-60"];
                    }
                    fcsNameTxt2.Color = Color.FromRgb(165, 82, 0);
                    fcsNameTxt2.Justify = AttachmentPoint.MiddleCenter;
                    fcsNameTxt2.TextString = unitD.XRs[i];
                    fcsNameTxt2.Height = 2.5;
                    fcsNameTxt2.Position = fcsPoly2.GetPoint3dAt(3).Add(new Vector3d(5, -2.06, 0));
                    fcsNameTxt2.HorizontalMode = TextHorizontalMode.TextCenter;
                    fcsNameTxt2.AlignmentPoint = fcsNameTxt2.Position;
                    modSpace.AppendEntity(fcsNameTxt2);
                    acTrans.AddNewlyCreatedDBObject(fcsNameTxt2, true);

                    if (i == 0) fcsPoint = fcsPoly2.GetPoint2dAt(3);
                    busPoint = fcsPoly1.GetPoint2dAt(0).Add(new Vector2d(0, -2.5));
                    FcsCount++;
                }

                if (busXR[busXR.Count - 1].BusPoints == null)
                {
                    Bus bXR = busXR[busXR.Count - 1];
                    bXR.BusPoints = new List<BusPoint>();
                    busXR[busXR.Count - 1] = bXR;
                    busXR[busXR.Count - 1].AddPointLastBusPoint(new Point3d(busPoint.X, busPoint.Y, 0));
                }
                else
                {
                    if (busXR[busXR.Count - 1].Term == unitD.XR && busXR[busXR.Count - 1].BusPoints[busXR[busXR.Count - 1].BusPoints.Count - 1].Page != true)
                    {
                        busXR[busXR.Count - 1].SetEndLastBusPoint(false);
                    }
                    busXR[busXR.Count - 1].AddPointLastBusPoint(new Point3d(busPoint.X, busPoint.Y, 0));
                }

                if (jumperPoints.Count >= 2)
                {
                    Polyline FcsJumperUp = new Polyline();
                    FcsJumperUp.SetDatabaseDefaults();
                    FcsJumperUp.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    FcsJumperUp.Closed = false;
                    FcsJumperUp.AddVertexAt(0, new Point2d(jumperPoints[0].X, jumperPoints[0].Y), 0, 0, 0);
                    FcsJumperUp.AddVertexAt(1, FcsJumperUp.GetPoint2dAt(0).Add(new Vector2d(0, 2.5)), 0, 0, 0);
                    FcsJumperUp.AddVertexAt(2, new Point2d(jumperPoints[1].X, jumperPoints[1].Y + 2.5), 0, 0, 0);
                    FcsJumperUp.AddVertexAt(3, FcsJumperUp.GetPoint2dAt(2).Add(new Vector2d(0, -2.5)), 0, 0, 0);
                    modSpace.AppendEntity(FcsJumperUp);
                    acTrans.AddNewlyCreatedDBObject(FcsJumperUp, true);

                    Polyline FcsJumperDown = new Polyline();
                    FcsJumperDown.SetDatabaseDefaults();
                    FcsJumperDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    FcsJumperDown.Closed = false;
                    FcsJumperDown.AddVertexAt(0, FcsJumperUp.GetPoint2dAt(0).Add(new Vector2d(0, -5 * (unitD.XRs.Count / 2))), 0, 0, 0);
                    FcsJumperDown.AddVertexAt(1, FcsJumperDown.GetPoint2dAt(0).Add(new Vector2d(0, -10)), 0, 0, 0);
                    FcsJumperDown.AddVertexAt(2, FcsJumperUp.GetPoint2dAt(3).Add(new Vector2d(0, -5 * (unitD.XRs.Count / 2) - 10)), 0, 0, 0);
                    FcsJumperDown.AddVertexAt(3, FcsJumperDown.GetPoint2dAt(2).Add(new Vector2d(0, 10)), 0, 0, 0);
                    modSpace.AppendEntity(FcsJumperDown);
                    acTrans.AddNewlyCreatedDBObject(FcsJumperDown, true);

                    DBText fcsNumTxt2 = new DBText();
                    fcsNumTxt2.SetDatabaseDefaults();
                    fcsNumTxt2.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        fcsNumTxt2.TextStyleId = tst["ROMANS0-60"];
                    }
                    fcsNumTxt2.Color = Color.FromRgb(165, 82, 0);
                    fcsNumTxt2.Justify = AttachmentPoint.MiddleCenter;
                    fcsNumTxt2.TextString = unitD.XRNum.ToString();
                    fcsNumTxt2.Height = 2.5;
                    fcsNumTxt2.Position = FcsJumperDown.GetPoint3dAt(2).Add(new Vector3d((FcsJumperDown.GetPoint3dAt(1).X - FcsJumperDown.GetPoint3dAt(2).X) / 2, 5, 0));
                    fcsNumTxt2.HorizontalMode = TextHorizontalMode.TextCenter;
                    fcsNumTxt2.AlignmentPoint = fcsNumTxt2.Position;
                    modSpace.AppendEntity(fcsNumTxt2);
                    acTrans.AddNewlyCreatedDBObject(fcsNumTxt2, true);
                }

                
                lineFromFcs.SetDatabaseDefaults();
                lineFromFcs.Color = Color.FromColorIndex(ColorMethod.ByLayer, 190);
                lineFromFcs.StartPoint = new Point3d(fcsPoint.X, fcsPoint.Y - 2.5, 0);
                lineFromFcs.EndPoint = lineFromFcs.StartPoint.Add(new Vector3d(-18, 0, 0));
                modSpace.AppendEntity(lineFromFcs);
                acTrans.AddNewlyCreatedDBObject(lineFromFcs, true);

                Point = lineFromFcs.EndPoint;
            }
            else
            {
                Point = point.Add(new Vector3d(-125, -20, 0));
            }

            Point2d xtPoint = Point2d.Origin;
            if (unitD.XTPins != null)
            {
                for (int i = 0; i < unitD.XTPins.Count; i++)
                {
                    Polyline xtPinPoly = new Polyline();
                    xtPinPoly.SetDatabaseDefaults();
                    xtPinPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    xtPinPoly.Closed = true;
                    xtPinPoly.AddVertexAt(0, new Point2d(Point.X, Point.Y + 2.5 - 5 * i), 0, 0, 0);
                    xtPinPoly.AddVertexAt(1, xtPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -5)), 0, 0, 0);
                    xtPinPoly.AddVertexAt(2, xtPinPoly.GetPoint2dAt(0).Add(new Vector2d(-10, -5)), 0, 0, 0);
                    xtPinPoly.AddVertexAt(3, xtPinPoly.GetPoint2dAt(0).Add(new Vector2d(-10, 0)), 0, 0, 0);
                    modSpace.AppendEntity(xtPinPoly);
                    acTrans.AddNewlyCreatedDBObject(xtPinPoly, true);

                    DBText xtPinNameTxt = new DBText();
                    xtPinNameTxt.SetDatabaseDefaults();
                    xtPinNameTxt.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        xtPinNameTxt.TextStyleId = tst["ROMANS0-60"];
                    }
                    xtPinNameTxt.Color = Color.FromRgb(165, 82, 0);
                    xtPinNameTxt.Justify = AttachmentPoint.MiddleCenter;
                    xtPinNameTxt.TextString = unitD.XTPins[i].ToString();
                    xtPinNameTxt.Height = 2.5;
                    xtPinNameTxt.Position = xtPinPoly.GetPoint3dAt(3).Add(new Vector3d(5, -2.06, 0));
                    xtPinNameTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                    xtPinNameTxt.AlignmentPoint = xtPinNameTxt.Position;
                    modSpace.AppendEntity(xtPinNameTxt);
                    acTrans.AddNewlyCreatedDBObject(xtPinNameTxt, true);

                    if (i < unitD.LinesNum)
                    {
                        Line lineFromXT = new Line();
                        lineFromXT.SetDatabaseDefaults();
                        lineFromXT.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                        lineFromXT.StartPoint = xtPinPoly.GetPoint3dAt(3).Add(new Vector3d(0, -2.5, 0));
                        lineFromXT.EndPoint = lineFromXT.StartPoint.Add(new Vector3d(-30, 0, 0));
                        modSpace.AppendEntity(lineFromXT);
                        acTrans.AddNewlyCreatedDBObject(lineFromXT, true);

                        group.Lines.Add(lineFromXT.EndPoint);
                    }

                    if (i == 0) 
                    { 
                        xtPoint = xtPinPoly.GetPoint2dAt(3);

                        if (!unitD.Table)
                        {
                            Line upLine1 = new Line();
                            upLine1.SetDatabaseDefaults();
                            upLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                            upLine1.StartPoint = xtPinPoly.GetPoint3dAt(0).Add(new Vector3d(0, -2.5, 0));
                            upLine1.EndPoint = upLine1.StartPoint.Add(new Vector3d(9.5, 4, 0));
                            modSpace.AppendEntity(upLine1);
                            acTrans.AddNewlyCreatedDBObject(upLine1, true);

                            Line upLine2 = new Line();
                            upLine2.SetDatabaseDefaults();
                            upLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                            upLine2.StartPoint = upLine1.EndPoint;
                            upLine2.EndPoint = upLine2.StartPoint.Add(new Vector3d(0, 21, 0));
                            modSpace.AppendEntity(upLine2);
                            acTrans.AddNewlyCreatedDBObject(upLine2, true);

                            DBText upLink = new DBText();
                            upLink.SetDatabaseDefaults();
                            upLink.WidthFactor = 0.5;
                            if (tst.Has("ROMANS0-60"))
                            {
                                upLink.TextStyleId = tst["ROMANS0-60"];
                            }
                            upLink.Color = Color.FromRgb(165, 82, 0);
                            upLink.Justify = AttachmentPoint.BottomRight;
                            upLink.TextString = makeLink24_7(int.Parse(unitD.XTPins[0].Replace("FU", "")), true, 1);
                            upLink.Rotation = 1.5708;
                            upLink.WidthFactor = 0.5;
                            upLink.Height = 3;
                            upLink.Position = upLine2.EndPoint.Add(new Vector3d(-1, -1, 0));
                            upLink.VerticalMode = TextVerticalMode.TextBottom;
                            upLink.AlignmentPoint = upLink.Position;
                            modSpace.AppendEntity(upLink);
                            acTrans.AddNewlyCreatedDBObject(upLink, true);

                            if (int.Parse(unitD.XTPins[0].Replace("FU", "")) < 10)
                            {
                                Line downLine1 = new Line();
                                downLine1.SetDatabaseDefaults();
                                downLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                                downLine1.StartPoint = xtPinPoly.GetPoint3dAt(0).Add(new Vector3d(0, -2.5, 0));
                                downLine1.EndPoint = downLine1.StartPoint.Add(new Vector3d(9.5, -4, 0));
                                modSpace.AppendEntity(downLine1);
                                acTrans.AddNewlyCreatedDBObject(downLine1, true);

                                Line downLine2 = new Line();
                                downLine2.SetDatabaseDefaults();
                                downLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                                downLine2.StartPoint = downLine1.EndPoint;
                                downLine2.EndPoint = downLine2.StartPoint.Add(new Vector3d(0, -21, 0));
                                modSpace.AppendEntity(downLine2);
                                acTrans.AddNewlyCreatedDBObject(downLine2, true);

                                DBText downLink = new DBText();
                                downLink.SetDatabaseDefaults();
                                downLink.WidthFactor = 0.5;
                                if (tst.Has("ROMANS0-60"))
                                {
                                    downLink.TextStyleId = tst["ROMANS0-60"];
                                }
                                downLink.Color = Color.FromRgb(165, 82, 0);
                                downLink.Justify = AttachmentPoint.BottomLeft;
                                downLink.TextString = makeLink24_7(int.Parse(unitD.XTPins[0].Replace("FU", "")), false, 1);
                                downLink.Rotation = 1.5708;
                                downLink.WidthFactor = 0.5;
                                downLink.Height = 3;
                                downLink.Position = downLine2.EndPoint.Add(new Vector3d(-1, 1, 0));
                                downLink.VerticalMode = TextVerticalMode.TextBottom;
                                downLink.AlignmentPoint = downLink.Position;
                                modSpace.AppendEntity(downLink);
                                acTrans.AddNewlyCreatedDBObject(downLink, true);
                            }
                        }
                    }
                    if (i == 1)
                    {
                        busPoint = xtPinPoly.GetPoint2dAt(0).Add(new Vector2d(0, -2.5));

                        if (!unitD.Table)
                        {
                            Line upLine1 = new Line();
                            upLine1.SetDatabaseDefaults();
                            upLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                            upLine1.StartPoint = xtPinPoly.GetPoint3dAt(0).Add(new Vector3d(0, -2.5, 0));
                            upLine1.EndPoint = upLine1.StartPoint.Add(new Vector3d(3, 2.5, 0));
                            modSpace.AppendEntity(upLine1);
                            acTrans.AddNewlyCreatedDBObject(upLine1, true);

                            Line upLine2 = new Line();
                            upLine2.SetDatabaseDefaults();
                            upLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                            upLine2.StartPoint = upLine1.EndPoint;
                            upLine2.EndPoint = upLine2.StartPoint.Add(new Vector3d(0, 25, 0));
                            modSpace.AppendEntity(upLine2);
                            acTrans.AddNewlyCreatedDBObject(upLine2, true);

                            DBText upLink = new DBText();
                            upLink.SetDatabaseDefaults();
                            upLink.WidthFactor = 0.5;
                            if (tst.Has("ROMANS0-60"))
                            {
                                upLink.TextStyleId = tst["ROMANS0-60"];
                            }
                            upLink.Color = Color.FromRgb(165, 82, 0);
                            upLink.Justify = AttachmentPoint.BottomRight;
                            upLink.TextString = makeLink24_7(int.Parse(unitD.XTPins[0].Replace("FU", "")), true, 1);
                            upLink.Rotation = 1.5708;
                            upLink.WidthFactor = 0.5;
                            upLink.Height = 3;
                            upLink.Position = upLine2.EndPoint.Add(new Vector3d(-1, -1, 0));
                            upLink.VerticalMode = TextVerticalMode.TextBottom;
                            upLink.AlignmentPoint = upLink.Position;
                            modSpace.AppendEntity(upLink);
                            acTrans.AddNewlyCreatedDBObject(upLink, true);

                            if (int.Parse(unitD.XTPins[0].Replace("FU", "")) < 10)
                            {
                                Line downLine1 = new Line();
                                downLine1.SetDatabaseDefaults();
                                downLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                                downLine1.StartPoint = xtPinPoly.GetPoint3dAt(0).Add(new Vector3d(0, -2.5, 0));
                                downLine1.EndPoint = downLine1.StartPoint.Add(new Vector3d(3, -2.5, 0));
                                modSpace.AppendEntity(downLine1);
                                acTrans.AddNewlyCreatedDBObject(downLine1, true);

                                Line downLine2 = new Line();
                                downLine2.SetDatabaseDefaults();
                                downLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                                downLine2.StartPoint = downLine1.EndPoint;
                                downLine2.EndPoint = downLine2.StartPoint.Add(new Vector3d(0, -25, 0));
                                modSpace.AppendEntity(downLine2);
                                acTrans.AddNewlyCreatedDBObject(downLine2, true);

                                DBText downLink = new DBText();
                                downLink.SetDatabaseDefaults();
                                downLink.WidthFactor = 0.5;
                                if (tst.Has("ROMANS0-60"))
                                {
                                    downLink.TextStyleId = tst["ROMANS0-60"];
                                }
                                downLink.Color = Color.FromRgb(165, 82, 0);
                                downLink.Justify = AttachmentPoint.BottomLeft;
                                downLink.TextString = makeLink24_7(int.Parse(unitD.XTPins[0].Replace("FU", "")), false, 1);
                                downLink.Rotation = 1.5708;
                                downLink.WidthFactor = 0.5;
                                downLink.Height = 3;
                                downLink.Position = downLine2.EndPoint.Add(new Vector3d(-1, 1, 0));
                                downLink.VerticalMode = TextVerticalMode.TextBottom;
                                downLink.AlignmentPoint = downLink.Position;
                                modSpace.AppendEntity(downLink);
                                acTrans.AddNewlyCreatedDBObject(downLink, true);
                            }
                        }
                    }
                }
                if (unitD.Table)
                    if (busXT[busXT.Count - 1].BusPoints == null)
                    {
                        Bus bXT = busXT[busXT.Count - 1];
                        bXT.BusPoints = new List<BusPoint>();
                        busXT[busXT.Count - 1] = bXT;
                        busXT[busXT.Count - 1].AddPointLastBusPoint(new Point3d(busPoint.X, busPoint.Y, 0));
                        busXT[busXT.Count - 1].SetPinFirstBusPoint(int.Parse(unitD.XTPins[1]));
                    }
                    else
                    {
                        if (busXT[busXT.Count - 1].Term == unitD.XT && busXT[busXT.Count - 1].BusPoints[busXT[busXT.Count - 1].BusPoints.Count - 1].Page != true)
                        {
                            busXT[busXT.Count - 1].SetEndLastBusPoint(false);
                        }
                        busXT[busXT.Count - 1].AddPointLastBusPoint(new Point3d(busPoint.X, busPoint.Y, 0));
                        busXT[busXT.Count - 1].SetPinLastBusPoint(int.Parse(unitD.XTPins[1]));
                    }
            }
            else
            {
                lineFromFcs.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                lineFromFcs.EndPoint = lineFromFcs.EndPoint.Add(new Vector3d(-30, 0, 0));
                group.Lines.Add(lineFromFcs.EndPoint);
            }

            if (terminal && unitD.Table)
            {
                Polyline xrPoly = new Polyline();
                xrPoly.SetDatabaseDefaults();
                xrPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                xrPoly.Closed = true;
                xrPoly.AddVertexAt(0, fcsPoint.Add(new Vector2d(0, 35)), 0, 0, 0);
                xrPoly.AddVertexAt(1, xrPoly.GetPoint2dAt(0).Add(new Vector2d(0, 10)), 0, 0, 0);
                xrPoly.AddVertexAt(2, xrPoly.GetPoint2dAt(0).Add(new Vector2d(25, 10)), 0, 0, 0);
                xrPoly.AddVertexAt(3, xrPoly.GetPoint2dAt(0).Add(new Vector2d(25, 0)), 0, 0, 0);
                modSpace.AppendEntity(xrPoly);
                acTrans.AddNewlyCreatedDBObject(xrPoly, true);

                DBText xrNameTxt = new DBText();
                xrNameTxt.SetDatabaseDefaults();
                xrNameTxt.WidthFactor = 0.5;
                if (tst.Has("ROMANS0-60"))
                {
                    xrNameTxt.TextStyleId = tst["ROMANS0-60"];
                }
                xrNameTxt.Color = Color.FromRgb(165, 82, 0);
                xrNameTxt.Justify = AttachmentPoint.MiddleCenter;
                xrNameTxt.TextString = unitD.XR;
                xrNameTxt.Height = 2.5;
                xrNameTxt.Position = xrPoly.GetPoint3dAt(1).Add(new Vector3d(10, -5, 0));
                xrNameTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                xrNameTxt.AlignmentPoint = xrNameTxt.Position;
                modSpace.AppendEntity(xrNameTxt);
                acTrans.AddNewlyCreatedDBObject(xrNameTxt, true);

                if (unitD.XTPins != null)
                {
                    Polyline xtPoly = new Polyline();
                    xtPoly.SetDatabaseDefaults();
                    xtPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    xtPoly.Closed = true;
                    xtPoly.AddVertexAt(0, xtPoint.Add(new Vector2d(-7.5, 35)), 0, 0, 0);
                    xtPoly.AddVertexAt(1, xtPoly.GetPoint2dAt(0).Add(new Vector2d(0, 10)), 0, 0, 0);
                    xtPoly.AddVertexAt(2, xtPoly.GetPoint2dAt(0).Add(new Vector2d(25, 10)), 0, 0, 0);
                    xtPoly.AddVertexAt(3, xtPoly.GetPoint2dAt(0).Add(new Vector2d(25, 0)), 0, 0, 0);
                    modSpace.AppendEntity(xtPoly);
                    acTrans.AddNewlyCreatedDBObject(xtPoly, true);

                    DBText xtNameTxt = new DBText();
                    xtNameTxt.SetDatabaseDefaults();
                    xtNameTxt.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        xtNameTxt.TextStyleId = tst["ROMANS0-60"];
                    }
                    xtNameTxt.Color = Color.FromRgb(165, 82, 0);
                    xtNameTxt.Justify = AttachmentPoint.MiddleCenter;
                    xtNameTxt.TextString = unitD.XT;
                    xtNameTxt.Height = 2.5;
                    xtNameTxt.Position = xtPoly.GetPoint3dAt(1).Add(new Vector3d(10, -5, 0));
                    xtNameTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                    xtNameTxt.AlignmentPoint = xtNameTxt.Position;
                    modSpace.AppendEntity(xtNameTxt);
                    acTrans.AddNewlyCreatedDBObject(xtNameTxt, true);
                }
            }
            group.UnitsD[unitIndex] = unitD;
            Groups[groupIndex] = group;
        }

        private static void DrawTableRow(Transaction acTrans, BlockTableRecord modSpace, Point3d point, string text)
        {
            Polyline poly = new Polyline();
            poly.SetDatabaseDefaults();
            poly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
            poly.Closed = true;
            poly.AddVertexAt(0, new Point2d(point.X, point.Y), 0, 0, 0);
            poly.AddVertexAt(1, new Point2d(point.X, point.Y - 4.12), 0, 0, 0);
            poly.AddVertexAt(2, new Point2d(point.X - 63, point.Y - 4.12), 0, 0, 0);
            poly.AddVertexAt(3, new Point2d(point.X - 63, point.Y), 0, 0, 0);
            modSpace.AppendEntity(poly);
            acTrans.AddNewlyCreatedDBObject(poly, true);

            DBText txt = new DBText();
            txt.SetDatabaseDefaults();
            txt.WidthFactor = 0.5;
            if (tst.Has("ROMANS0-60"))
            {
                txt.TextStyleId = tst["ROMANS0-60"];
            }
            txt.Color = Color.FromRgb(165, 82, 0);
            txt.Justify = AttachmentPoint.MiddleLeft;
            txt.TextString = text;
            txt.Height = 2.5;
            txt.Position = poly.GetPoint3dAt(3).Add(new Vector3d(1, -2.06, 0));
            txt.HorizontalMode = TextHorizontalMode.TextLeft;
            txt.AlignmentPoint = txt.Position;
            modSpace.AppendEntity(txt);
            acTrans.AddNewlyCreatedDBObject(txt, true);
        }

        private static void DrawBus(Transaction acTrans, BlockTableRecord modSpace)
        {
            #region XR
            for (int i = 0; i < busXR.Count; i++)
            {
                for (int j=0; j<busXR[i].BusPoints.Count; j++)
                {
                    if (j == 0)
                    {
                        Line lineRightOblique = new Line();
                        lineRightOblique.SetDatabaseDefaults();
                        lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightOblique.StartPoint = busXR[i].BusPoints[j].Point;
                        lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, 2, 0));
                        modSpace.AppendEntity(lineRightOblique);
                        acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                        Line lineRightStraight = new Line();
                        lineRightStraight.SetDatabaseDefaults();
                        lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                        lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, 35, 0));
                        modSpace.AppendEntity(lineRightStraight);
                        acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                        DBText busRightTxt = new DBText();
                        busRightTxt.SetDatabaseDefaults();
                        busRightTxt.WidthFactor = 0.5;
                        if (tst.Has("ROMANS0-60"))
                        {
                            busRightTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        busRightTxt.Color = Color.FromRgb(165, 82, 0);
                        busRightTxt.Justify = AttachmentPoint.BottomRight;
                        busRightTxt.TextString = "Шина FBST 500-PLCGY";
                        busRightTxt.Rotation = 1.5708;
                        busRightTxt.WidthFactor = 0.5;
                        busRightTxt.Height = 3;
                        busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, -1, 0));
                        busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                        busRightTxt.AlignmentPoint = busRightTxt.Position;
                        modSpace.AppendEntity(busRightTxt);
                        acTrans.AddNewlyCreatedDBObject(busRightTxt, true);

                        Line lineLeftOblique = new Line();
                        lineLeftOblique.SetDatabaseDefaults();
                        lineLeftOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftOblique.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                        lineLeftOblique.EndPoint = lineLeftOblique.StartPoint.Add(new Vector3d(-3, 2, 0));
                        modSpace.AppendEntity(lineLeftOblique);
                        acTrans.AddNewlyCreatedDBObject(lineLeftOblique, true);

                        Line lineLeftStraight = new Line();
                        lineLeftStraight.SetDatabaseDefaults();
                        lineLeftStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftStraight.StartPoint = lineLeftOblique.EndPoint;
                        lineLeftStraight.EndPoint = lineLeftStraight.StartPoint.Add(new Vector3d(0, 35, 0));
                        modSpace.AppendEntity(lineLeftStraight);
                        acTrans.AddNewlyCreatedDBObject(lineLeftStraight, true);

                        DBText busLeftTxt = new DBText();
                        busLeftTxt.SetDatabaseDefaults();
                        busLeftTxt.WidthFactor = 0.5;
                        if (tst.Has("ROMANS0-60"))
                        {
                            busLeftTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        busLeftTxt.Color = Color.FromRgb(165, 82, 0);
                        busLeftTxt.Justify = AttachmentPoint.BottomRight;
                        busLeftTxt.TextString = "Шина FBST 500-PLCGY";
                        busLeftTxt.Rotation = 1.5708;
                        busLeftTxt.WidthFactor = 0.5;
                        busLeftTxt.Height = 3;
                        busLeftTxt.Position = lineLeftStraight.EndPoint.Add(new Vector3d(-1, -1, 0));
                        busLeftTxt.VerticalMode = TextVerticalMode.TextBottom;
                        busLeftTxt.AlignmentPoint = busLeftTxt.Position;
                        modSpace.AppendEntity(busLeftTxt);
                        acTrans.AddNewlyCreatedDBObject(busLeftTxt, true);

                        if (busXR[i].BusPoints.Count==2)
                        {
                            Line lineRightUpOblique1 = new Line();
                            lineRightUpOblique1.SetDatabaseDefaults();
                            lineRightUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightUpOblique1.StartPoint = busXR[i].BusPoints[j].Point;
                            lineRightUpOblique1.EndPoint = lineRightUpOblique1.StartPoint.Add(new Vector3d(3, -2, 0));
                            modSpace.AppendEntity(lineRightUpOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineRightUpOblique1, true);

                            Line lineRightDownOblique1 = new Line();
                            lineRightDownOblique1.SetDatabaseDefaults();
                            lineRightDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightDownOblique1.StartPoint = busXR[i].BusPoints[j+1].Point;
                            lineRightDownOblique1.EndPoint = lineRightDownOblique1.StartPoint.Add(new Vector3d(3, 2, 0));
                            modSpace.AppendEntity(lineRightDownOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineRightDownOblique1, true);

                            Line lineRightStraight1 = new Line();
                            lineRightStraight1.SetDatabaseDefaults();
                            lineRightStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightStraight1.StartPoint = lineRightUpOblique1.EndPoint;
                            lineRightStraight1.EndPoint = lineRightDownOblique1.EndPoint;
                            modSpace.AppendEntity(lineRightStraight1);
                            acTrans.AddNewlyCreatedDBObject(lineRightStraight1, true);

                            Line lineLeftUpOblique1 = new Line();
                            lineLeftUpOblique1.SetDatabaseDefaults();
                            lineLeftUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineLeftUpOblique1.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                            lineLeftUpOblique1.EndPoint = lineLeftUpOblique1.StartPoint.Add(new Vector3d(-3, -2, 0));
                            modSpace.AppendEntity(lineLeftUpOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineLeftUpOblique1, true);

                            Line lineLeftDownOblique1 = new Line();
                            lineLeftDownOblique1.SetDatabaseDefaults();
                            lineLeftDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineLeftDownOblique1.StartPoint = busXR[i].BusPoints[j+1].Point.Add(new Vector3d(-25, 0, 0));
                            lineLeftDownOblique1.EndPoint = lineLeftDownOblique1.StartPoint.Add(new Vector3d(-3, 2, 0));
                            modSpace.AppendEntity(lineLeftDownOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineLeftDownOblique1, true);

                            Line lineLeftStraight1 = new Line();
                            lineLeftStraight1.SetDatabaseDefaults();
                            lineLeftStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineLeftStraight1.StartPoint = lineLeftUpOblique1.EndPoint;
                            lineLeftStraight1.EndPoint = lineLeftDownOblique1.EndPoint;
                            modSpace.AppendEntity(lineLeftStraight1);
                            acTrans.AddNewlyCreatedDBObject(lineLeftStraight1, true);
                        }
                    }
                    else if (j==busXR[i].BusPoints.Count - 1)
                    {
                        if (!busXR[i].BusPoints[j].End)
                        {
                            Line lineRightOblique = new Line();
                            lineRightOblique.SetDatabaseDefaults();
                            lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightOblique.StartPoint = busXR[i].BusPoints[j].Point;
                            lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, -2, 0));
                            modSpace.AppendEntity(lineRightOblique);
                            acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                            Line lineRightStraight = new Line();
                            lineRightStraight.SetDatabaseDefaults();
                            lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                            lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                            modSpace.AppendEntity(lineRightStraight);
                            acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                            DBText busRightTxt = new DBText();
                            busRightTxt.SetDatabaseDefaults();
                            busRightTxt.WidthFactor = 0.5;
                            if (tst.Has("ROMANS0-60"))
                            {
                                busRightTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            busRightTxt.Color = Color.FromRgb(165, 82, 0);
                            busRightTxt.Justify = AttachmentPoint.BottomRight;
                            busRightTxt.TextString = "Шина FBST 500-PLCGY";
                            busRightTxt.Rotation = 1.5708;
                            busRightTxt.WidthFactor = 0.5;
                            busRightTxt.Height = 3;
                            busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                            busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                            busRightTxt.AlignmentPoint = busRightTxt.Position;
                            modSpace.AppendEntity(busRightTxt);
                            acTrans.AddNewlyCreatedDBObject(busRightTxt, true);

                            Line lineLeftOblique = new Line();
                            lineLeftOblique.SetDatabaseDefaults();
                            lineLeftOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineLeftOblique.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                            lineLeftOblique.EndPoint = lineLeftOblique.StartPoint.Add(new Vector3d(-3, -2, 0));
                            modSpace.AppendEntity(lineLeftOblique);
                            acTrans.AddNewlyCreatedDBObject(lineLeftOblique, true);

                            Line lineLeftStraight = new Line();
                            lineLeftStraight.SetDatabaseDefaults();
                            lineLeftStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineLeftStraight.StartPoint = lineLeftOblique.EndPoint;
                            lineLeftStraight.EndPoint = lineLeftStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                            modSpace.AppendEntity(lineLeftStraight);
                            acTrans.AddNewlyCreatedDBObject(lineLeftStraight, true);

                            DBText busLeftTxt = new DBText();
                            busLeftTxt.SetDatabaseDefaults();
                            busLeftTxt.WidthFactor = 0.5;
                            if (tst.Has("ROMANS0-60"))
                            {
                                busLeftTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            busLeftTxt.Color = Color.FromRgb(165, 82, 0);
                            busLeftTxt.Justify = AttachmentPoint.BottomRight;
                            busLeftTxt.TextString = "Шина FBST 500-PLCGY";
                            busLeftTxt.Rotation = 1.5708;
                            busLeftTxt.WidthFactor = 0.5;
                            busLeftTxt.Height = 3;
                            busLeftTxt.Position = lineLeftStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                            busLeftTxt.VerticalMode = TextVerticalMode.TextBottom;
                            busLeftTxt.AlignmentPoint = busLeftTxt.Position;
                            modSpace.AppendEntity(busLeftTxt);
                            acTrans.AddNewlyCreatedDBObject(busLeftTxt, true);
                        }
                        else if (busXR[i].BusPoints[j].Page)
                        {
                            if (i < busXR.Count - 1)
                            {
                                if (busXR[i].Term == busXR[i + 1].Term)
                                {
                                    Line lineRightOblique = new Line();
                                    lineRightOblique.SetDatabaseDefaults();
                                    lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                    lineRightOblique.StartPoint = busXR[i].BusPoints[j].Point;
                                    lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, -2, 0));
                                    modSpace.AppendEntity(lineRightOblique);
                                    acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                                    Line lineRightStraight = new Line();
                                    lineRightStraight.SetDatabaseDefaults();
                                    lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                    lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                                    lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                                    modSpace.AppendEntity(lineRightStraight);
                                    acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                                    DBText busRightTxt = new DBText();
                                    busRightTxt.SetDatabaseDefaults();
                                    busRightTxt.WidthFactor = 0.5;
                                    busRightTxt.Color = Color.FromRgb(165, 82, 0);
                                    busRightTxt.Justify = AttachmentPoint.BottomRight;
                                    busRightTxt.TextString = "Шина FBST 500-PLCGY";
                                    busRightTxt.Rotation = 1.5708;
                                    busRightTxt.WidthFactor = 0.5;
                                    if (tst.Has("ROMANS0-60"))
                                    {
                                        busRightTxt.TextStyleId = tst["ROMANS0-60"];
                                    }
                                    busRightTxt.Height = 3;
                                    busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                                    busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                                    busRightTxt.AlignmentPoint = busRightTxt.Position;
                                    modSpace.AppendEntity(busRightTxt);
                                    acTrans.AddNewlyCreatedDBObject(busRightTxt, true);

                                    Line lineLeftOblique = new Line();
                                    lineLeftOblique.SetDatabaseDefaults();
                                    lineLeftOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                    lineLeftOblique.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                                    lineLeftOblique.EndPoint = lineLeftOblique.StartPoint.Add(new Vector3d(-3, -2, 0));
                                    modSpace.AppendEntity(lineLeftOblique);
                                    acTrans.AddNewlyCreatedDBObject(lineLeftOblique, true);

                                    Line lineLeftStraight = new Line();
                                    lineLeftStraight.SetDatabaseDefaults();
                                    lineLeftStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                    lineLeftStraight.StartPoint = lineLeftOblique.EndPoint;
                                    lineLeftStraight.EndPoint = lineLeftStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                                    modSpace.AppendEntity(lineLeftStraight);
                                    acTrans.AddNewlyCreatedDBObject(lineLeftStraight, true);

                                    DBText busLeftTxt = new DBText();
                                    busLeftTxt.SetDatabaseDefaults();
                                    busLeftTxt.WidthFactor = 0.5;
                                    if (tst.Has("ROMANS0-60"))
                                    {
                                        busLeftTxt.TextStyleId = tst["ROMANS0-60"];
                                    }
                                    busLeftTxt.Color = Color.FromRgb(165, 82, 0);
                                    busLeftTxt.Justify = AttachmentPoint.BottomRight;
                                    busLeftTxt.TextString = "Шина FBST 500-PLCGY";
                                    busLeftTxt.Rotation = 1.5708;
                                    busLeftTxt.WidthFactor = 0.5;
                                    busLeftTxt.Height = 3;
                                    busLeftTxt.Position = lineLeftStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                                    busLeftTxt.VerticalMode = TextVerticalMode.TextBottom;
                                    busLeftTxt.AlignmentPoint = busLeftTxt.Position;
                                    modSpace.AppendEntity(busLeftTxt);
                                    acTrans.AddNewlyCreatedDBObject(busLeftTxt, true);
                                }
                            }
                        }
                    }
                    else
                    {
                        Line lineRightUpOblique1 = new Line();
                        lineRightUpOblique1.SetDatabaseDefaults();
                        lineRightUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightUpOblique1.StartPoint = busXR[i].BusPoints[j-1].Point;
                        lineRightUpOblique1.EndPoint = lineRightUpOblique1.StartPoint.Add(new Vector3d(3, -2, 0));
                        modSpace.AppendEntity(lineRightUpOblique1);
                        acTrans.AddNewlyCreatedDBObject(lineRightUpOblique1, true);

                        Line lineRightDownOblique1 = new Line();
                        lineRightDownOblique1.SetDatabaseDefaults();
                        lineRightDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightDownOblique1.StartPoint = busXR[i].BusPoints[j].Point;
                        lineRightDownOblique1.EndPoint = lineRightDownOblique1.StartPoint.Add(new Vector3d(3, 2, 0));
                        modSpace.AppendEntity(lineRightDownOblique1);
                        acTrans.AddNewlyCreatedDBObject(lineRightDownOblique1, true);

                        Line lineRightStraight1 = new Line();
                        lineRightStraight1.SetDatabaseDefaults();
                        lineRightStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightStraight1.StartPoint = lineRightUpOblique1.EndPoint;
                        lineRightStraight1.EndPoint = lineRightDownOblique1.EndPoint;
                        modSpace.AppendEntity(lineRightStraight1);
                        acTrans.AddNewlyCreatedDBObject(lineRightStraight1, true); 
                        
                        Line lineLeftUpOblique1 = new Line();
                        lineLeftUpOblique1.SetDatabaseDefaults();
                        lineLeftUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftUpOblique1.StartPoint = busXR[i].BusPoints[j-1].Point.Add(new Vector3d(-25, 0, 0));
                        lineLeftUpOblique1.EndPoint = lineLeftUpOblique1.StartPoint.Add(new Vector3d(-3, -2, 0));
                        modSpace.AppendEntity(lineLeftUpOblique1);
                        acTrans.AddNewlyCreatedDBObject(lineLeftUpOblique1, true);

                        Line lineLeftDownOblique1 = new Line();
                        lineLeftDownOblique1.SetDatabaseDefaults();
                        lineLeftDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftDownOblique1.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                        lineLeftDownOblique1.EndPoint = lineLeftDownOblique1.StartPoint.Add(new Vector3d(-3, 2, 0));
                        modSpace.AppendEntity(lineLeftDownOblique1);
                        acTrans.AddNewlyCreatedDBObject(lineLeftDownOblique1, true);

                        Line lineLeftStraight1 = new Line();
                        lineLeftStraight1.SetDatabaseDefaults();
                        lineLeftStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftStraight1.StartPoint = lineLeftUpOblique1.EndPoint;
                        lineLeftStraight1.EndPoint = lineLeftDownOblique1.EndPoint;
                        modSpace.AppendEntity(lineLeftStraight1);
                        acTrans.AddNewlyCreatedDBObject(lineLeftStraight1, true);

                        Line lineRightUpOblique2 = new Line();
                        lineRightUpOblique2.SetDatabaseDefaults();
                        lineRightUpOblique2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightUpOblique2.StartPoint = busXR[i].BusPoints[j].Point;
                        lineRightUpOblique2.EndPoint = lineRightUpOblique2.StartPoint.Add(new Vector3d(3, -2, 0));
                        modSpace.AppendEntity(lineRightUpOblique2);
                        acTrans.AddNewlyCreatedDBObject(lineRightUpOblique2, true);

                        Line lineRightDownOblique2 = new Line();
                        lineRightDownOblique2.SetDatabaseDefaults();
                        lineRightDownOblique2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightDownOblique2.StartPoint = busXR[i].BusPoints[j + 1].Point;
                        lineRightDownOblique2.EndPoint = lineRightDownOblique2.StartPoint.Add(new Vector3d(3, 2, 0));
                        modSpace.AppendEntity(lineRightDownOblique2);
                        acTrans.AddNewlyCreatedDBObject(lineRightDownOblique2, true);

                        Line lineRightStraight = new Line();
                        lineRightStraight.SetDatabaseDefaults();
                        lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineRightStraight.StartPoint = lineRightUpOblique2.EndPoint;
                        lineRightStraight.EndPoint = lineRightDownOblique2.EndPoint;
                        modSpace.AppendEntity(lineRightStraight);
                        acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                        Line lineLeftUpOblique = new Line();
                        lineLeftUpOblique.SetDatabaseDefaults();
                        lineLeftUpOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftUpOblique.StartPoint = busXR[i].BusPoints[j].Point.Add(new Vector3d(-25, 0, 0));
                        lineLeftUpOblique.EndPoint = lineLeftUpOblique.StartPoint.Add(new Vector3d(-3, -2, 0));
                        modSpace.AppendEntity(lineLeftUpOblique);
                        acTrans.AddNewlyCreatedDBObject(lineLeftUpOblique, true);

                        Line lineLeftDownOblique = new Line();
                        lineLeftDownOblique.SetDatabaseDefaults();
                        lineLeftDownOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftDownOblique.StartPoint = busXR[i].BusPoints[j+1].Point.Add(new Vector3d(-25, 0, 0));
                        lineLeftDownOblique.EndPoint = lineLeftDownOblique.StartPoint.Add(new Vector3d(-3, 2, 0));
                        modSpace.AppendEntity(lineLeftDownOblique);
                        acTrans.AddNewlyCreatedDBObject(lineLeftDownOblique, true);

                        Line lineLeftStraight = new Line();
                        lineLeftStraight.SetDatabaseDefaults();
                        lineLeftStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        lineLeftStraight.StartPoint = lineLeftUpOblique.EndPoint;
                        lineLeftStraight.EndPoint = lineLeftDownOblique.EndPoint;
                        modSpace.AppendEntity(lineLeftStraight);
                        acTrans.AddNewlyCreatedDBObject(lineLeftStraight, true); 
                    }
                }
            }
            #endregion
            #region XT
            for (int i = 0; i < busXT.Count; i++)
            {
                if (busXT[i].BusPoints != null)
                {
                    for (int j = 0; j < busXT[i].BusPoints.Count; j++)
                    {
                        if (j == 0)
                        {
                            Line lineRightOblique = new Line();
                            lineRightOblique.SetDatabaseDefaults();
                            lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightOblique.StartPoint = busXT[i].BusPoints[j].Point;
                            lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, 2, 0));
                            modSpace.AppendEntity(lineRightOblique);
                            acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                            Line lineRightStraight = new Line();
                            lineRightStraight.SetDatabaseDefaults();
                            lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                            lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, 35, 0));
                            modSpace.AppendEntity(lineRightStraight);
                            acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                            DBText busRightTxt = new DBText();
                            busRightTxt.SetDatabaseDefaults();
                            busRightTxt.WidthFactor = 0.5;
                            if (tst.Has("ROMANS0-60"))
                            {
                                busRightTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            busRightTxt.Color = Color.FromRgb(165, 82, 0);
                            busRightTxt.Justify = AttachmentPoint.BottomRight;
                            string link = busXT[i].Type == "DI" ? makeLinkDI(busXT[i].BusPoints[j].Pin, true) : makeLinkDO(busXT[i].BusPoints[j].Pin, true);
                            busRightTxt.TextString = link != "0" ? "В " + busXT[i].Term + ":" + link : busXT[i].Link;
                            busRightTxt.Rotation = 1.5708;
                            busRightTxt.WidthFactor = 0.5;
                            busRightTxt.Height = 3;
                            busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, -1, 0));
                            busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                            busRightTxt.AlignmentPoint = busRightTxt.Position;
                            modSpace.AppendEntity(busRightTxt);
                            acTrans.AddNewlyCreatedDBObject(busRightTxt, true);

                            if (busXT[i].BusPoints.Count == 2)
                            {
                                Line lineRightUpOblique1 = new Line();
                                lineRightUpOblique1.SetDatabaseDefaults();
                                lineRightUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                lineRightUpOblique1.StartPoint = busXT[i].BusPoints[j].Point;
                                lineRightUpOblique1.EndPoint = lineRightUpOblique1.StartPoint.Add(new Vector3d(3, -2, 0));
                                modSpace.AppendEntity(lineRightUpOblique1);
                                acTrans.AddNewlyCreatedDBObject(lineRightUpOblique1, true);

                                Line lineRightDownOblique1 = new Line();
                                lineRightDownOblique1.SetDatabaseDefaults();
                                lineRightDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                lineRightDownOblique1.StartPoint = busXT[i].BusPoints[j + 1].Point;
                                lineRightDownOblique1.EndPoint = lineRightDownOblique1.StartPoint.Add(new Vector3d(3, 2, 0));
                                modSpace.AppendEntity(lineRightDownOblique1);
                                acTrans.AddNewlyCreatedDBObject(lineRightDownOblique1, true);

                                Line lineRightStraight1 = new Line();
                                lineRightStraight1.SetDatabaseDefaults();
                                lineRightStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                lineRightStraight1.StartPoint = lineRightUpOblique1.EndPoint;
                                lineRightStraight1.EndPoint = lineRightDownOblique1.EndPoint;
                                modSpace.AppendEntity(lineRightStraight1);
                                acTrans.AddNewlyCreatedDBObject(lineRightStraight1, true);
                            }
                        }
                        else if (j == busXT[i].BusPoints.Count - 1)
                        {
                            if (!busXT[i].BusPoints[j].End)
                            {
                                Line lineRightOblique = new Line();
                                lineRightOblique.SetDatabaseDefaults();
                                lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                lineRightOblique.StartPoint = busXT[i].BusPoints[j].Point;
                                lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, -2, 0));
                                modSpace.AppendEntity(lineRightOblique);
                                acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                                Line lineRightStraight = new Line();
                                lineRightStraight.SetDatabaseDefaults();
                                lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                                lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                                modSpace.AppendEntity(lineRightStraight);
                                acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                                DBText busRightTxt = new DBText();
                                busRightTxt.SetDatabaseDefaults();
                                busRightTxt.WidthFactor = 0.5;
                                if (tst.Has("ROMANS0-60"))
                                {
                                    busRightTxt.TextStyleId = tst["ROMANS0-60"];
                                }
                                busRightTxt.Color = Color.FromRgb(165, 82, 0);
                                busRightTxt.Justify = AttachmentPoint.BottomRight;
                                string link = busXT[i].Type == "DI" ? makeLinkDI(busXT[i].BusPoints[j].Pin, false) : makeLinkDO(busXT[i].BusPoints[j].Pin, false);
                                busRightTxt.TextString = link != "0" ? "В " + busXT[i].Term + ":" + link : busXT[i].Link;
                                busRightTxt.Rotation = 1.5708;
                                busRightTxt.WidthFactor = 0.5;
                                busRightTxt.Height = 3;
                                busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                                busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                                busRightTxt.AlignmentPoint = busRightTxt.Position;
                                modSpace.AppendEntity(busRightTxt);
                                acTrans.AddNewlyCreatedDBObject(busRightTxt, true);
                            }
                            else if (busXT[i].BusPoints[j].Page)
                            {
                                if (i < busXT.Count - 1)
                                {
                                    if (busXT[i].Term == busXT[i + 1].Term)
                                    {
                                        Line lineRightOblique = new Line();
                                        lineRightOblique.SetDatabaseDefaults();
                                        lineRightOblique.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                        lineRightOblique.StartPoint = busXT[i].BusPoints[j].Point;
                                        lineRightOblique.EndPoint = lineRightOblique.StartPoint.Add(new Vector3d(3, -2, 0));
                                        modSpace.AppendEntity(lineRightOblique);
                                        acTrans.AddNewlyCreatedDBObject(lineRightOblique, true);

                                        Line lineRightStraight = new Line();
                                        lineRightStraight.SetDatabaseDefaults();
                                        lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                                        lineRightStraight.StartPoint = lineRightOblique.EndPoint;
                                        lineRightStraight.EndPoint = lineRightStraight.StartPoint.Add(new Vector3d(0, -35, 0));
                                        modSpace.AppendEntity(lineRightStraight);
                                        acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);

                                        DBText busRightTxt = new DBText();
                                        busRightTxt.SetDatabaseDefaults();
                                        busRightTxt.WidthFactor = 0.5;
                                        if (tst.Has("ROMANS0-60"))
                                        {
                                            busRightTxt.TextStyleId = tst["ROMANS0-60"];
                                        }
                                        busRightTxt.Color = Color.FromRgb(165, 82, 0);
                                        busRightTxt.Justify = AttachmentPoint.BottomRight;
                                        string link = busXT[i].Type == "DI" ? makeLinkDI(busXT[i].BusPoints[j].Pin, false) : makeLinkDO(busXT[i].BusPoints[j].Pin, false);
                                        busRightTxt.TextString = link != "0" ? "В " + busXT[i].Term + ":" + link : busXT[i].Link;
                                        busRightTxt.Rotation = 1.5708;
                                        busRightTxt.WidthFactor = 0.5;
                                        busRightTxt.Height = 3;
                                        busRightTxt.Position = lineRightStraight.EndPoint.Add(new Vector3d(-1, 1, 0));
                                        busRightTxt.VerticalMode = TextVerticalMode.TextBottom;
                                        busRightTxt.AlignmentPoint = busRightTxt.Position;
                                        modSpace.AppendEntity(busRightTxt);
                                        acTrans.AddNewlyCreatedDBObject(busRightTxt, true);
                                    }
                                }
                            }
                        }
                        else
                        {
                            Line lineRightUpOblique1 = new Line();
                            lineRightUpOblique1.SetDatabaseDefaults();
                            lineRightUpOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightUpOblique1.StartPoint = busXT[i].BusPoints[j - 1].Point;
                            lineRightUpOblique1.EndPoint = lineRightUpOblique1.StartPoint.Add(new Vector3d(3, -2, 0));
                            modSpace.AppendEntity(lineRightUpOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineRightUpOblique1, true);

                            Line lineRightDownOblique1 = new Line();
                            lineRightDownOblique1.SetDatabaseDefaults();
                            lineRightDownOblique1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightDownOblique1.StartPoint = busXT[i].BusPoints[j].Point;
                            lineRightDownOblique1.EndPoint = lineRightDownOblique1.StartPoint.Add(new Vector3d(3, 2, 0));
                            modSpace.AppendEntity(lineRightDownOblique1);
                            acTrans.AddNewlyCreatedDBObject(lineRightDownOblique1, true);

                            Line lineRightStraight1 = new Line();
                            lineRightStraight1.SetDatabaseDefaults();
                            lineRightStraight1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightStraight1.StartPoint = lineRightUpOblique1.EndPoint;
                            lineRightStraight1.EndPoint = lineRightDownOblique1.EndPoint;
                            modSpace.AppendEntity(lineRightStraight1);
                            acTrans.AddNewlyCreatedDBObject(lineRightStraight1, true);

                            Line lineRightUpOblique2 = new Line();
                            lineRightUpOblique2.SetDatabaseDefaults();
                            lineRightUpOblique2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightUpOblique2.StartPoint = busXT[i].BusPoints[j].Point;
                            lineRightUpOblique2.EndPoint = lineRightUpOblique2.StartPoint.Add(new Vector3d(3, -2, 0));
                            modSpace.AppendEntity(lineRightUpOblique2);
                            acTrans.AddNewlyCreatedDBObject(lineRightUpOblique2, true);

                            Line lineRightDownOblique2 = new Line();
                            lineRightDownOblique2.SetDatabaseDefaults();
                            lineRightDownOblique2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightDownOblique2.StartPoint = busXT[i].BusPoints[j + 1].Point;
                            lineRightDownOblique2.EndPoint = lineRightDownOblique2.StartPoint.Add(new Vector3d(3, 2, 0));
                            modSpace.AppendEntity(lineRightDownOblique2);
                            acTrans.AddNewlyCreatedDBObject(lineRightDownOblique2, true);

                            Line lineRightStraight = new Line();
                            lineRightStraight.SetDatabaseDefaults();
                            lineRightStraight.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            lineRightStraight.StartPoint = lineRightUpOblique2.EndPoint;
                            lineRightStraight.EndPoint = lineRightDownOblique2.EndPoint;
                            modSpace.AppendEntity(lineRightStraight);
                            acTrans.AddNewlyCreatedDBObject(lineRightStraight, true);
                        }
                    }
                }
            }
            #endregion
        }

        private static void DrawConnectionSchemePart(Transaction acTrans, BlockTableRecord modSpace)
        {
            for (int i = 0; i < Groups.Count; i++)
            {
                if (Groups[i].Lines.Count > 1)
                {
                    Line xtJumper = new Line();
                    xtJumper.SetDatabaseDefaults();
                    xtJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 5);
                    xtJumper.StartPoint = Groups[i].Lines[0];
                    xtJumper.EndPoint = Groups[i].Lines[Groups[i].Lines.Count - 1];
                    modSpace.AppendEntity(xtJumper);
                    acTrans.AddNewlyCreatedDBObject(xtJumper, true);

                    double y = xtJumper.EndPoint.Y + (xtJumper.StartPoint.Y - xtJumper.EndPoint.Y) / 2;
                    double x = xtJumper.StartPoint.X - 180;

                    if (Groups[i].unit != null)
                    {
                        MText designationTxt = new MText();
                        designationTxt.SetDatabaseDefaults();
                        designationTxt.TextHeight = 3;
                        designationTxt.Width = 50;
                        designationTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 3);
                        designationTxt.Location = new Point3d(x, y, 0);
                        designationTxt.Contents = Groups[i].unit == null ? Groups[i].ccunit.Value.designation : Groups[i].unit.Value.param;
                        designationTxt.Attachment = AttachmentPoint.MiddleCenter;
                        if (tst.Has("ROMANS0-60"))
                        {
                            designationTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        modSpace.AppendEntity(designationTxt);
                        acTrans.AddNewlyCreatedDBObject(designationTxt, true);

                        Polyline designationPoly = new Polyline();
                        designationPoly.SetDatabaseDefaults();
                        designationPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        designationPoly.Closed = true;
                        designationPoly.AddVertexAt(0, new Point2d(x + 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                        designationPoly.AddVertexAt(1, new Point2d(x - 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                        designationPoly.AddVertexAt(2, new Point2d(x - 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                        designationPoly.AddVertexAt(3, new Point2d(x + 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                        modSpace.AppendEntity(designationPoly);
                        acTrans.AddNewlyCreatedDBObject(designationPoly, true);
                    }
                    else
                    {
                        Line cableFromXtJumper = new Line();
                        cableFromXtJumper.SetDatabaseDefaults();
                        cableFromXtJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 5);
                        cableFromXtJumper.StartPoint = new Point3d(xtJumper.StartPoint.X, y, 0);
                        cableFromXtJumper.EndPoint = cableFromXtJumper.StartPoint.Add(new Vector3d(-60, 0, 0));
                        modSpace.AppendEntity(cableFromXtJumper);
                        acTrans.AddNewlyCreatedDBObject(cableFromXtJumper, true);

                        Line xtJumperLeft = new Line();
                        xtJumperLeft.SetDatabaseDefaults();
                        xtJumperLeft.Color = Color.FromColorIndex(ColorMethod.ByLayer, 5);
                        xtJumperLeft.StartPoint = cableFromXtJumper.EndPoint.Add(new Vector3d(0, 2.5 * (Groups[i].ccunit.Value.terminalsOut.Count - 1), 0));
                        xtJumperLeft.EndPoint = xtJumperLeft.StartPoint.Add(new Vector3d(0, -5 * (Groups[i].ccunit.Value.terminalsOut.Count - 1), 0));
                        modSpace.AppendEntity(xtJumperLeft);
                        acTrans.AddNewlyCreatedDBObject(xtJumperLeft, true);

                        Point3d point = xtJumperLeft.StartPoint;

                        MText termNameTxt = new MText();
                        termNameTxt.SetDatabaseDefaults();
                        termNameTxt.TextHeight = 3;
                        termNameTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 1);
                        termNameTxt.Location = point.Add(new Vector3d(-43, 3, 0));
                        termNameTxt.Contents = Groups[i].ccunit.Value.terminalsOut[0].Split(':')[0];
                        termNameTxt.Attachment = AttachmentPoint.BottomLeft;
                        if (tst.Has("ROMANS0-60"))
                        {
                            termNameTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        modSpace.AppendEntity(termNameTxt);
                        acTrans.AddNewlyCreatedDBObject(termNameTxt, true);

                        Polyline cupNamePoly = new Polyline();
                        cupNamePoly.SetDatabaseDefaults();
                        cupNamePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 3);
                        cupNamePoly.Closed = true;
                        cupNamePoly.AddVertexAt(0, new Point2d(cableFromXtJumper.EndPoint.X + 4, cableFromXtJumper.EndPoint.Y + 12), 0, 0, 0);
                        cupNamePoly.AddVertexAt(1, cupNamePoly.GetPoint2dAt(0).Add(new Vector2d(-55, 0)), 0, 0, 0);
                        cupNamePoly.AddVertexAt(2, cupNamePoly.GetPoint2dAt(1).Add(new Vector2d(0, -24)), 0, 0, 0);
                        cupNamePoly.AddVertexAt(3, cupNamePoly.GetPoint2dAt(2).Add(new Vector2d(55, 0)), 0, 0, 0);
                        if (lineTypeTable.Has("штриховая2"))
                            cupNamePoly.LinetypeId = lineTypeTable["штриховая2"];
                        else if (lineTypeTable.Has("hidden2"))
                            cupNamePoly.LinetypeId = lineTypeTable["hidden2"];
                        cupNamePoly.LinetypeScale = 15;
                        modSpace.AppendEntity(cupNamePoly);
                        acTrans.AddNewlyCreatedDBObject(cupNamePoly, true);

                        MText cupNameTxt = new MText();
                        cupNameTxt.SetDatabaseDefaults();
                        cupNameTxt.TextHeight = 3;
                        cupNameTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 32);
                        cupNameTxt.Location = cupNamePoly.GetPoint3dAt(1).Add(new Vector3d(1, 0, 0));
                        string str = CupboardNum < 10 ? "0" + CupboardNum : CupboardNum.ToString();
                        cupNameTxt.Contents = "Шкаф VA-V" + str;
                        cupNameTxt.Attachment = AttachmentPoint.BottomLeft;
                        if (tst.Has("ROMANS0-60"))
                        {
                            cupNameTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        modSpace.AppendEntity(cupNameTxt);
                        acTrans.AddNewlyCreatedDBObject(cupNameTxt, true);

                        for (int j = 0; j < Groups[i].Lines.Count; j++)
                        {
                            MText termLinkTxt = new MText();
                            termLinkTxt.SetDatabaseDefaults();
                            termLinkTxt.TextHeight = 3;
                            termLinkTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 32);
                            termLinkTxt.Location = Groups[i].Lines[j].Add(new Vector3d(1, 0, 0));
                            str = CupboardNum < 10 ? "0" + CupboardNum : CupboardNum.ToString();
                            termLinkTxt.Contents = "VA-V" + str + "/" + (int.Parse(Groups[i].ccunit.Value.param) * 2 - (Groups[i].ccunit.Value.terminalsOut.Count - j)).ToString();
                            termLinkTxt.Attachment = AttachmentPoint.BottomLeft;
                            if (tst.Has("ROMANS0-60"))
                            {
                                termLinkTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            modSpace.AppendEntity(termLinkTxt);
                            acTrans.AddNewlyCreatedDBObject(termLinkTxt, true);
                        }

                        for (int j = 0; j < Groups[i].ccunit.Value.terminalsOut.Count; j++)
                        {
                            Line lineFromJumper = new Line();
                            lineFromJumper.SetDatabaseDefaults();
                            lineFromJumper.Color = Color.FromColorIndex(ColorMethod.ByLayer, 161);
                            lineFromJumper.StartPoint = point;
                            lineFromJumper.EndPoint = lineFromJumper.StartPoint.Add(new Vector3d(-30, 0, 0));
                            modSpace.AppendEntity(lineFromJumper);
                            acTrans.AddNewlyCreatedDBObject(lineFromJumper, true);

                            MText termLinkTxt = new MText();
                            termLinkTxt.SetDatabaseDefaults();
                            termLinkTxt.TextHeight = 3;
                            termLinkTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 32);
                            termLinkTxt.Location = lineFromJumper.EndPoint.Add(new Vector3d(1, 0, 0));
                            str = CupboardNum < 10 ? "0" + CupboardNum : CupboardNum.ToString();
                            termLinkTxt.Contents = "VA-V" + str + "/" + (int.Parse(Groups[i].ccunit.Value.param) * 2 - (Groups[i].ccunit.Value.terminalsOut.Count - j)).ToString();
                            termLinkTxt.Attachment = AttachmentPoint.BottomLeft;
                            if (tst.Has("ROMANS0-60"))
                            {
                                termLinkTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            modSpace.AppendEntity(termLinkTxt);
                            acTrans.AddNewlyCreatedDBObject(termLinkTxt, true);

                            Polyline termPoly = new Polyline();
                            termPoly.SetDatabaseDefaults();
                            termPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                            termPoly.Closed = true;
                            termPoly.AddVertexAt(0, new Point2d(lineFromJumper.EndPoint.X, lineFromJumper.EndPoint.Y + 2.5), 0, 0, 0);
                            termPoly.AddVertexAt(1, termPoly.GetPoint2dAt(0).Add(new Vector2d(-13, 0)), 0, 0, 0);
                            termPoly.AddVertexAt(2, termPoly.GetPoint2dAt(1).Add(new Vector2d(0, -5)), 0, 0, 0);
                            termPoly.AddVertexAt(3, termPoly.GetPoint2dAt(2).Add(new Vector2d(13, 0)), 0, 0, 0);
                            modSpace.AppendEntity(termPoly);
                            acTrans.AddNewlyCreatedDBObject(termPoly, true);

                            MText termTxt = new MText();
                            termTxt.SetDatabaseDefaults();
                            termTxt.TextHeight = 3;
                            termTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 32);
                            termTxt.Location = termPoly.GetPoint3dAt(2).Add(new Vector3d(6.5, 2.5, 0));
                            termTxt.Contents = Groups[i].ccunit.Value.terminalsOut[j].Contains(":") ? Groups[i].ccunit.Value.terminalsOut[j].Split(':')[1] : Groups[i].ccunit.Value.terminalsOut[j];
                            termTxt.Attachment = AttachmentPoint.MiddleCenter;
                            if (tst.Has("ROMANS0-60"))
                            {
                                termTxt.TextStyleId = tst["ROMANS0-60"];
                            }
                            modSpace.AppendEntity(termTxt);
                            acTrans.AddNewlyCreatedDBObject(termTxt, true);

                            point = point.Add(new Vector3d(0, -5, 0));
                        }
                        

                        Line cableNameLine = new Line();
                        cableNameLine.SetDatabaseDefaults();
                        cableNameLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        cableNameLine.StartPoint = cableFromXtJumper.StartPoint.Add(new Vector3d(-30, 0, 0));
                        cableNameLine.EndPoint = cableNameLine.StartPoint.Add(new Vector3d(0, 20, 0));
                        modSpace.AppendEntity(cableNameLine);
                        acTrans.AddNewlyCreatedDBObject(cableNameLine, true);

                        Polyline cableNamePoly = new Polyline();
                        cableNamePoly.SetDatabaseDefaults();
                        cableNamePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        cableNamePoly.Closed = true;
                        cableNamePoly.AddVertexAt(0, new Point2d(cableNameLine.EndPoint.X + 4, cableNameLine.EndPoint.Y), 0, 0, 0);
                        cableNamePoly.AddVertexAt(1, cableNamePoly.GetPoint2dAt(0).Add(new Vector2d(0, 5)), 0, 0, 0);
                        cableNamePoly.AddVertexAt(2, cableNamePoly.GetPoint2dAt(1).Add(new Vector2d(-35, 0)), 0, 0, 0);
                        cableNamePoly.AddVertexAt(3, cableNamePoly.GetPoint2dAt(2).Add(new Vector2d(0, -5)), 0, 0, 0);
                        modSpace.AppendEntity(cableNamePoly);
                        acTrans.AddNewlyCreatedDBObject(cableNamePoly, true);

                        MText cableNameTxt = new MText();
                        cableNameTxt.SetDatabaseDefaults();
                        cableNameTxt.TextHeight = 3;
                        cableNameTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        cableNameTxt.Location = cableNamePoly.GetPoint3dAt(3).Add(new Vector3d(1, 0, 0));
                        string cNum = CupboardNum<10?"0"+CupboardNum:CupboardNum.ToString();
                        cableNameTxt.Contents = "VA-V" + cNum + "/HPL-0" + cNum + "-" + Groups[i].ccunit.Value.param;
                        cableNameTxt.Attachment = AttachmentPoint.BottomLeft;
                        if (tst.Has("ROMANS0-60"))
                        {
                            cableNameTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        modSpace.AppendEntity(cableNameTxt);
                        acTrans.AddNewlyCreatedDBObject(cableNameTxt, true);

                        MText designationTxt = new MText();
                        designationTxt.SetDatabaseDefaults();
                        designationTxt.TextHeight = 3;
                        designationTxt.Width = 50;
                        designationTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 3);
                        designationTxt.Location = new Point3d(x, y, 0);
                        designationTxt.Contents = Groups[i].unit == null ? Groups[i].ccunit.Value.designation : Groups[i].unit.Value.param;
                        designationTxt.Attachment = AttachmentPoint.MiddleCenter;
                        if (tst.Has("ROMANS0-60"))
                        {
                            designationTxt.TextStyleId = tst["ROMANS0-60"];
                        }
                        modSpace.AppendEntity(designationTxt);
                        acTrans.AddNewlyCreatedDBObject(designationTxt, true);

                        Polyline designationPoly = new Polyline();
                        designationPoly.SetDatabaseDefaults();
                        designationPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                        designationPoly.Closed = true;
                        designationPoly.AddVertexAt(0, new Point2d(x + 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                        designationPoly.AddVertexAt(1, new Point2d(x - 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                        designationPoly.AddVertexAt(2, new Point2d(x - 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                        designationPoly.AddVertexAt(3, new Point2d(x + 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                        modSpace.AppendEntity(designationPoly);
                        acTrans.AddNewlyCreatedDBObject(designationPoly, true);
                    }
                }
                else
                {
                    double y = Groups[i].Lines[0].Y;
                    double x = Groups[i].Lines[0].X - 120;

                    DBText link = new DBText();
                    link.SetDatabaseDefaults();
                    link.WidthFactor = 0.5;
                    if (tst.Has("ROMANS0-60"))
                    {
                        link.TextStyleId = tst["ROMANS0-60"];
                    }
                    link.Color = Color.FromRgb(165, 82, 0);
                    link.Justify = AttachmentPoint.BottomLeft;
                    link.TextString = Groups[i].UnitsD[0].XTLink;
                    link.WidthFactor = 0.5;
                    link.Height = 3;
                    link.Position = Groups[i].Lines[0].Add(new Vector3d(2, 0, 0));
                    link.HorizontalMode = TextHorizontalMode.TextLeft;
                    link.AlignmentPoint = link.Position;
                    modSpace.AppendEntity(link);
                    acTrans.AddNewlyCreatedDBObject(link, true);

                    MText designationTxt = new MText();
                    designationTxt.SetDatabaseDefaults();
                    designationTxt.TextHeight = 3;
                    designationTxt.Width = 50;
                    designationTxt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 3);
                    designationTxt.Location = new Point3d(x, y, 0);
                    designationTxt.Contents = Groups[i].UnitsD[0].Equipment;
                    designationTxt.Attachment = AttachmentPoint.MiddleCenter;
                    if (tst.Has("ROMANS0-60"))
                    {
                        designationTxt.TextStyleId = tst["ROMANS0-60"];
                    }
                    modSpace.AppendEntity(designationTxt);
                    acTrans.AddNewlyCreatedDBObject(designationTxt, true);

                    Polyline designationPoly = new Polyline();
                    designationPoly.SetDatabaseDefaults();
                    designationPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 6);
                    designationPoly.Closed = true;
                    designationPoly.AddVertexAt(0, new Point2d(x + 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                    designationPoly.AddVertexAt(1, new Point2d(x - 27, y + designationTxt.ActualHeight / 2 + 5), 0, 0, 0);
                    designationPoly.AddVertexAt(2, new Point2d(x - 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                    designationPoly.AddVertexAt(3, new Point2d(x + 27, y - designationTxt.ActualHeight / 2 - 5), 0, 0, 0);
                    modSpace.AppendEntity(designationPoly);
                    acTrans.AddNewlyCreatedDBObject(designationPoly, true);
                }
            }
        }

        private static string makeLinkDI(int pin, bool up)
        {
            if (pin > 3)
            {
                if (up) return (pin - 3).ToString();
                else return (pin + 3).ToString();
            }
            else return "0";
        }
        private static string makeLinkDO(int pin, bool up)
        {
            if (pin >= 2)
            {
                if (up) return (pin - 1).ToString();
                else return (pin + 1).ToString();
            }
            else return "0";
        }
        private static string makeLink24_7(int pin, bool up, int number)
        {
            if (number==1)
            {
                if (pin == 1)
                {
                    if (up) return "В -T1:4L-5";
                    else return "В -ХТ24.7:" + (pin + 1);
                }
                else if (up) return "В -ХТ24.7:" + (pin - 1);
                else return "В -ХТ24.7:" + (pin + 1);
            }
            else
            {
                if (pin == 1)
                {
                    if (up) return "В -T1:4L-4";
                    else return "В -ХТ24.7:" + (pin + 1);
                }
                else if (up) return "В -ХТ24.7:" + (pin - 1);
                else return "В -ХТ24.7:" + (pin + 1);
            }
        }
        private static List<Unit?> loadDataFromConnectionScheme(string path)
        {
            List<Unit?> units = new List<Unit?>();
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

        //private static List<CCUnit?> loadDataFromCupboardConnections(string path)
        //{
        //    List<CCUnit?> units = new List<CCUnit?>();
        //    DataSet dataSet = new DataSet("EXCEL");
        //    string connectionString;
        //    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;MAXSCANROWS=0'";
        //    OleDbConnection connection = new OleDbConnection(connectionString);
        //    connection.Open();

        //    System.Data.DataTable schemaTable = connection.GetSchema("Tables", new string[] { null, null, null, "TABLE" });
        //    string sheet1 = (string)schemaTable.Rows[1].ItemArray[2];

        //    string select = String.Format("SELECT * FROM ["+sheet1+"]");
        //    OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
        //    adapter.FillSchema(dataSet, SchemaType.Source);
        //    adapter.Fill(dataSet);
        //    connection.Close();

        //    for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
        //    {
        //        if (dataSet.Tables[0].Rows[row][1].ToString() != "")
        //        {
        //            CCUnit unit = new CCUnit();
        //            unit.designation = dataSet.Tables[0].Rows[row][1].ToString();
        //            unit.terminalsIn = new List<string>();
        //            unit.terminalsIn.Add(dataSet.Tables[0].Rows[row][4].ToString());
        //            unit.terminalsIn.Add(dataSet.Tables[0].Rows[row][5].ToString());
        //            unit.terminalsOut = new List<string>();
        //            unit.terminalsOut.Add(dataSet.Tables[0].Rows[row][6].ToString());
        //            unit.terminalsOut.Add(dataSet.Tables[0].Rows[row][7].ToString());
        //            unit.param = dataSet.Tables[0].Rows[row][23].ToString();
        //            units.Add(unit);
        //        }
        //    }
        //    return units;
        //}

        private static List<CCUnit?> loadDataFromCupboardConnections(string path)
        {
            List<CCUnit?> units = new List<CCUnit?>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetSchema("Tables", new string[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [" + sheet1 + "]");
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.FillSchema(dataSet, SchemaType.Source);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 0; row < dataSet.Tables[0].Rows.Count; row+=2)
            {
                if (dataSet.Tables[0].Rows[row][1].ToString() != "")
                {
                    CCUnit unit = new CCUnit();
                    unit.designation = dataSet.Tables[0].Rows[row][1].ToString();
                    unit.terminalsOut = new List<string>();
                    unit.terminalsOut.Add(dataSet.Tables[0].Rows[row][3].ToString());
                    unit.terminalsOut.Add(dataSet.Tables[0].Rows[row+1][3].ToString());

                    unit.terminalsIn = new List<string>();
                    unit.terminalsIn.Add(dataSet.Tables[0].Rows[row][6].ToString());
                    unit.terminalsIn.Add(dataSet.Tables[0].Rows[row+1][6].ToString());

                    unit.param = dataSet.Tables[0].Rows[row][5].ToString();
                    units.Add(unit);
                }
            }
            return units;
        }

        private static List<Group> FindGroups()
        {
            List<Group> groups = new List<Group>();
            for (int i = 0; i < Units.Count; i++)
            {
                if (Units[i].Value.param != "Резерв" && Units[i].Value.param != "резерв" && !Units[i].Value.cupboardName.Contains("VA"))
                {
                    int step = 0;
                    if (Units[i].Value.terminals.Count % 3 != 0) step = 2;
                    else
                    {
                        bool exist = false;
                        for (int k = 0; k < Units[i].Value.terminals.Count; k++)
                            if (Units[i].Value.terminals[k].Contains("XT24.7")) exist = true;
                        if (exist) step = 2;
                        else
                        {
                            string prevTerm = "";
                            int prevPin = 0;
                            List<int> list = new List<int>();
                            for (int k = 0; k < Units[i].Value.terminals.Count; k++)
                            {
                                if (prevTerm == "")
                                {
                                    prevTerm = Units[i].Value.terminals[k].Split('[', ']')[1];
                                    int.TryParse(Units[i].Value.terminals[k].Split('[', ']')[2], out prevPin);
                                    list.Add(1);
                                }
                                else
                                {
                                    if (prevTerm == Units[i].Value.terminals[k].Split('[', ']')[1])
                                    {
                                        int pin = 0;
                                        int.TryParse(Units[i].Value.terminals[k].Split('[', ']')[2], out pin);
                                        if (prevPin + 1 == pin) { list[list.Count - 1]++; prevPin = pin; }
                                        else
                                        {
                                            int pin1 = 0;
                                            int.TryParse(Units[i].Value.terminals[k - 1].Split('[', ']')[2].Replace("FU", ""), out pin1);
                                            if (pin == pin1) { list[list.Count - 1]++; prevPin = pin; }
                                            else
                                            {
                                                list.Add(1);
                                                prevPin = pin;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int pin = 0;
                                        int.TryParse(Units[i].Value.terminals[k].Split('[', ']')[2], out pin);
                                        list.Add(1);
                                        prevPin = pin;
                                        prevTerm = Units[i].Value.terminals[k].Split('[', ']')[1];
                                    }
                                }
                            }
                            if (list.Count == 1 && list[0] == Units[i].Value.terminals.Count) step = 3;
                            else
                            {
                                step = Units[i].Value.terminals.Count / list.Count;
                            }
                        }
                    }
                    
                    int tryParse = 0;
                    List<Terminal> Terminals = new List<Terminal>();
                    for (int k = 0; k < Units[i].Value.terminals.Count; k += step)
                    {
                        int curStep = k;
                        bool addedTerm = false;
                        List<string> terminals = new List<string>();
                        while (curStep < k + step)
                        {
                            string termTag = Units[i].Value.terminals[curStep].Split('[', ']')[1];
                            string terminal = Units[i].Value.terminals[curStep].Split('[', ']')[2];
                            terminals.Add(terminal);
                            curStep++;
                        }
                        tryParse = 0;
                        Units[i].Value.terminals[k] = Transliteration.Front(Units[i].Value.terminals[k]);
                        if (Units[i].Value.terminals[k].Contains("XT4") || Units[i].Value.terminals[k].Contains("XT5"))
                        {
                            if (terminals.Count == 2)
                            {
                                if (int.TryParse(terminals[0], out tryParse))
                                    if (int.TryParse(terminals[1], out tryParse))
                                    {
                                        tryParse++;
                                        terminals.Add(tryParse.ToString());
                                        addedTerm = true;
                                    }
                            }
                        }
                        Terminal term = new Terminal();
                        term.LinesNum = addedTerm ? terminals.Count - 1 : terminals.Count;
                        term.Name = Units[i].Value.terminals[k].Split('[', ']')[1];
                        term.Terminals = new List<string>();
                        term.Terminals = terminals;
                        Terminals.Add(term);
                    }
                    Group group = new Group();
                    for (int k = 0; k < Terminals.Count; k++)
                    {
                        for (int j = 0; j < UnitsA.Count; j++)
                        {
                            string unitATerm = UnitsA[j].Terminals.Replace("-", "");
                            if (unitATerm == Terminals[k].Name && UnitsA[j].Fu.SequenceEqual(Terminals[k].Terminals))
                            {
                                if (group.UnitsA == null) { group.UnitsA = new List<UnitA>(); group.UnitsA.Add(UnitsA[j]); group.unit = Units[i]; UnitsA.RemoveAt(j); break; }
                                else { group.UnitsA.Add(UnitsA[j]); group.unit = Units[i]; UnitsA.RemoveAt(j); break; }
                            }
                        }
                    }
                    for (int k = 0; k < Terminals.Count; k++)
                    {
                        for (int j = 0; j < UnitsD.Count; j++)
                        {
                            string unitDTerm = UnitsD[j].XT.Replace("-", "");
                            if (unitDTerm == Terminals[k].Name && UnitsD[j].XTPins.SequenceEqual(Terminals[k].Terminals))
                            {
                                if (group.UnitsD == null) 
                                { 
                                    group.UnitsD = new List<UnitD>();
                                    UnitD unitD = UnitsD[j];
                                    unitD.LinesNum = Terminals[k].LinesNum;
                                    UnitsD[j] = unitD; 
                                    group.UnitsD.Add(UnitsD[j]); 
                                    group.unit = Units[i]; 
                                    UnitsD.RemoveAt(j); 
                                    break; 
                                }
                                else 
                                {
                                    UnitD unitD = UnitsD[j];
                                    unitD.LinesNum = Terminals[k].LinesNum;
                                    UnitsD[j] = unitD; 
                                    group.UnitsD.Add(UnitsD[j]); 
                                    group.unit = Units[i]; 
                                    UnitsD.RemoveAt(j); 
                                    break; 
                                }
                            }
                        }
                    }
                    
                    if (group.UnitsA != null || group.UnitsD != null) groups.Add(group);
                }
            }
            for (int j = 0; j < UnitsD.Count; j++)
            {
                if (UnitsD[j].XTPins==null)
                {
                    Group group = new Group();
                    group.UnitsD = new List<UnitD>();
                    group.UnitsD.Add(UnitsD[j]);
                    UnitsD.RemoveAt(j);
                    j--;
                    groups.Add(group);
                }
            }
            if (UnitsD.Count>0)
            {
                for (int i = 0; i < CCUnits.Count; i++)
                {
                    bool added = false;
                    List<string> terminals = new List<string>();
                    for (int k = 0; k < CCUnits[i].Value.terminalsIn.Count; k++)
                    {
                        CCUnits[i].Value.terminalsIn[k] = Transliteration.Front(CCUnits[i].Value.terminalsIn[k]);
                        string termTag = CCUnits[i].Value.terminalsIn[k].Split('-')[0];
                        string terminal = CCUnits[i].Value.terminalsIn[k].Split('-')[1];
                        terminals.Add(terminal);
                    }
                    int tryParse = 0;
                    if (int.TryParse(terminals[0], out tryParse))
                    {
                        int.TryParse(terminals[terminals.Count - 1], out tryParse);
                        tryParse++;
                        terminals.Add(tryParse.ToString());
                        added = true;
                    }
                    Terminal Term = new Terminal();
                    Term.LinesNum = added ? terminals.Count - 1 : terminals.Count;
                    Term.Name = CCUnits[i].Value.terminalsIn[0].Split('-')[0];
                    Term.Terminals = new List<string>();
                    Term.Terminals = terminals;

                    Group group = new Group();
                    for (int j = 0; j < UnitsD.Count; j++)
                    {
                        string unitDTerm = UnitsD[j].XT.Replace("-", "");
                        if (unitDTerm == Term.Name && UnitsD[j].XTPins.SequenceEqual(Term.Terminals))
                        {
                                group.UnitsD = new List<UnitD>();
                                UnitD unitD = UnitsD[j];
                                unitD.LinesNum = Term.LinesNum;
                                UnitsD[j] = unitD;
                                group.UnitsD.Add(UnitsD[j]);
                                group.ccunit = CCUnits[i];
                                group.unit = null;
                                groups.Add(group);
                                UnitsD.RemoveAt(j);
                                break;
                        }
                    }
                }
            }
            return groups;
        }

        struct Terminal
        {
            public string Name;
            public List<string> Terminals;
            public int LinesNum;
        }

        public static void ParseControlScheme()
        {
            List<UnitA> unitsA = new List<UnitA>();
            List<UnitD> unitsD = new List<UnitD>();
             
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;
            Point3d currentPoint;
            int fuCount = 0;

            PromptPointResult point = editor.GetPoint("Выберите левый верхний угол модуля DO на 8 странице: ");
            if (point.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else currentPoint = point.Value;

            PromptIntegerResult fu = editor.GetInteger("\nВведите количество листов с FU: ");
            if (fu.Status != PromptStatus.OK)
            {
                editor.WriteMessage("Неверный ввод...");
                return;
            }
            else fuCount = fu.Value;

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.SelectCrossingWindow(currentPoint, currentPoint.Add(new Vector3d(594, -415, 0)), filter);
            if (selRes.Status == PromptStatus.OK)
            {
                ObjectId obj = selRes.Value.GetObjectIds().ToList()[0];
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                    {
                        Table table = (Table)acTrans.GetObject(obj, OpenMode.ForRead);
                        if (table != null)
                        {
                            for (int j = 0; j < table.Rows.Count; j++)
                            {
                                if (!table.Cells[j, 0].TextString.Contains("Резерв") && !table.Cells[j, 0].TextString.Contains("резерв") && table.Cells[j, 0].TextString != "")
                                {
                                    UnitA unit = new UnitA();
                                    unit.Terminals = "-XT24.7";
                                    unit.Equipment = table.Cells[j, 0].GetTextString(FormatOption.FormatOptionNone);
                                    unit.Type = "DO";
                                    unit.Equipment = unit.Equipment.Replace("\n", "");
                                    unit.Fu = new List<string>();
                                    unit.Fu.Add("FU" + (j + 1));
                                    unit.Fu.Add((j + 1).ToString());
                                    unitsA.Add(unit);
                                }
                            }
                        }

                        currentPoint = currentPoint.Add(new Vector3d(420+329, 0, 0));
                        int currentRack = 3;
                        int currentSlot = 4;
                        int currentPin = 2;
                        int pages = 0;
                        int currentChannel = 0;

                        filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
                        filter = new SelectionFilter(filterlist);
                        selRes = editor.SelectCrossingWindow(currentPoint.Add(new Vector3d(0, 0, 0)), currentPoint.Add(new Vector3d(794, -415, 0)), filter);
                        if (selRes.Status == PromptStatus.OK)
                        {
                            obj = selRes.Value.GetObjectIds().ToList()[0];
                            table = (Table)acTrans.GetObject(obj, OpenMode.ForRead);
                            if (table != null)
                            {
                                for (int j = 0; j < table.Rows.Count; j += 8)
                                {
                                    for (int k = 0; k < 8; k++)
                                    {
                                        if (!table.Cells[j + k, 0].TextString.Contains("Резерв") && !table.Cells[j + k, 0].TextString.Contains("резерв") && table.Cells[j + k, 0].TextString != "")
                                        {
                                            UnitA unit = new UnitA();
                                            unit.Article = "6ES7331-7HF01-0AB0";
                                            unit.Equipment = table.Cells[j + k, 0].GetTextString(FormatOption.FormatOptionNone);
                                            unit.Equipment = unit.Equipment.Replace("\n", "");
                                            unit.Rack = currentRack;
                                            unit.Slot = currentSlot;
                                            unit.Terminals = "-XT" + currentRack + "." + currentSlot;
                                            unit.Fu = new List<string>();
                                            unit.Fu.Add("FU" + (k + 1));
                                            unit.Fu.Add((k + 1).ToString());
                                            unit.Type = "AI";
                                            unit.Pins = new List<string>();
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Channel = currentChannel;
                                            currentChannel++;
                                            unitsA.Add(unit);
                                        }
                                        else { currentPin += 2; currentChannel++; }
                                        if (k == 3) currentPin += 2;
                                    }
                                    currentSlot++;
                                    pages++;
                                    currentPin = 2;
                                    currentChannel = 0;
                                }
                            }
                        }

                        currentPoint = currentPoint.Add(new Vector3d(420 * pages + 329, 0, 0));
                        currentPin = 3;
                        pages = 0;
                        currentRack = 3;
                        currentChannel = 0;

                        filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
                        filter = new SelectionFilter(filterlist);
                        selRes = editor.SelectCrossingWindow(currentPoint.Add(new Vector3d(0, 0, 0)), currentPoint.Add(new Vector3d(794, -415, 0)), filter);
                        if (selRes.Status == PromptStatus.OK)
                        {
                            obj = selRes.Value.GetObjectIds().ToList()[0];
                            table = (Table)acTrans.GetObject(obj, OpenMode.ForRead);
                            if (table != null)
                            {
                                for (int j = 0; j < table.Rows.Count; j += 8)
                                {
                                    for (int k = 0; k < 8; k++)
                                    {
                                        if (!table.Cells[j + k, 0].TextString.Contains("Резерв") && !table.Cells[j + k, 0].TextString.Contains("резерв") && table.Cells[j + k, 0].TextString != "")
                                        {
                                            UnitA unit = new UnitA();
                                            unit.Article = "6ES7332-5HF00-0AB0";
                                            unit.Terminals = "-XT" + currentRack + "." + currentSlot;
                                            unit.Equipment = table.Cells[j + k, 0].GetTextString(FormatOption.FormatOptionNone);
                                            unit.Equipment = unit.Equipment.Replace("\n", "");
                                            unit.Rack = currentRack;
                                            unit.Slot = currentSlot;
                                            unit.Fu = new List<string>();
                                            unit.Fu.Add("FU" + (k + 1));
                                            unit.Fu.Add((k + 1).ToString());
                                            unit.Type = "AO";
                                            unit.Pins = new List<string>();
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Pins.Add(currentPin.ToString());
                                            currentPin++;
                                            unit.Channel = currentChannel;
                                            currentChannel++;
                                            unitsA.Add(unit);
                                        }
                                        else { currentPin += 4; currentChannel++; }
                                        if (k == 3) currentPin += 5;
                                    }
                                    currentSlot++;
                                    pages++;
                                    currentPin = 3;
                                    currentChannel = 0;
                                }
                            }
                        }


                        currentPoint = currentPoint.Add(new Vector3d(420 * pages + 329, 0, 0));
                        currentPin = 2;
                        int currentXTPin = 1;
                        pages = 0;
                        currentRack = 4;
                        currentSlot = 4;
                        currentChannel = 0;
                        int currentPart = 1;
                        int currentModNum = 1;
                        int currentFCNum = 1;
                        bool firstPage = true;
                        int downPinPart = 1;
                        int downPinNum = 0;
                        int currentFu = 11;

                        filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
                        filter = new SelectionFilter(filterlist);
                        selRes = editor.SelectCrossingWindow(currentPoint.Add(new Vector3d(0, 0, 0)), currentPoint.Add(new Vector3d(794, -415, 0)), filter);
                        if (selRes.Status == PromptStatus.OK)
                        {
                            obj = selRes.Value.GetObjectIds().ToList()[0];
                            table = (Table)acTrans.GetObject(obj, OpenMode.ForRead);
                            if (table != null)
                            {
                                for (int j = 0; j < table.Rows.Count; j += 16)
                                {
                                    for (int k = 0; k < 16; k++)
                                    {
                                        if (!table.Cells[j + k, 0].TextString.Contains("Резерв") && !table.Cells[j + k, 0].TextString.Contains("резерв") && table.Cells[j + k, 0].TextString != "")
                                        {
                                            UnitD unit = new UnitD();
                                            unit.Article = "6ES7321-1BL00-0AA0";
                                            unit.XT = "-XT" + currentRack + "." + currentSlot + "." + currentPart;
                                            unit.Equipment = table.Cells[j + k, 0].GetTextString(FormatOption.FormatOptionNone);
                                            unit.Equipment = unit.Equipment.Replace("\n", "");
                                            unit.Rack = currentRack;
                                            unit.Slot = currentSlot;
                                            unit.Type = "DI";
                                            unit.XR = "-" + currentPart + "XRI" + currentRack + "." + currentSlot;
                                            unit.LeftPin = currentPin;
                                            currentPin++;
                                            unit.RightPin = "IN" + downPinPart + "." + downPinNum;
                                            downPinNum++;
                                            unit.XTLink = "-XT24.2:FU" + currentFu + "/3.3";
                                            if (!firstPage)
                                            {
                                                unit.XTPins = new List<string>();
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                            }
                                            else
                                            {
                                                unit.XTLink = currentRack + "." + currentSlot + "." + currentPart + "-" + currentFCNum;
                                            }
                                            unit.Channel = currentChannel;
                                            currentChannel++;

                                            unit.XRNum = currentFCNum;
                                            currentFCNum++;
                                            unit.XRs = new List<string>();
                                            unit.XRs.Add("A1");
                                            unit.XRs.Add("14");
                                            unit.XRs.Add("A2");
                                            unit.XRs.Add("11");

                                            unitsD.Add(unit);
                                        }
                                        else { currentPin++; downPinNum++; currentChannel++; currentXTPin += 3; currentFCNum++; }
                                        if (k == 7) { currentPin += 2; downPinNum = 0; downPinPart++; }
                                    }
                                    if (currentPart == 1) { currentPart++; currentPin += 2; downPinNum = 0; downPinPart = 3; currentXTPin = 1; }
                                    else
                                    {
                                        currentPart = 1;
                                        currentSlot++;
                                        currentPin = 2;
                                        currentModNum++;
                                        downPinPart = 1;
                                        downPinNum = 0;
                                        currentXTPin = 1;
                                        currentFu++;
                                        if (currentModNum == 10)
                                        {
                                            currentModNum = 1;
                                            currentRack++;
                                        }
                                        currentChannel = 0;
                                    }
                                    pages++;
                                    firstPage = false;
                                    currentFCNum = 1;
                                }
                            }
                        }

                        currentPoint = currentPoint.Add(new Vector3d(420 * pages + 329, 0, 0));
                        currentPin = 2;
                        currentXTPin = 1;
                        pages = 0;
                        currentChannel = 0;
                        currentPart = 1;
                        currentFCNum = 1;
                        downPinPart = 1;
                        downPinNum = 0;
                        firstPage = true;

                        filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
                        filter = new SelectionFilter(filterlist);
                        selRes = editor.SelectCrossingWindow(currentPoint.Add(new Vector3d(0, 0, 0)), currentPoint.Add(new Vector3d(794, -415, 0)), filter);
                        if (selRes.Status == PromptStatus.OK)
                        {
                            obj = selRes.Value.GetObjectIds().ToList()[0];
                            table = (Table)acTrans.GetObject(obj, OpenMode.ForRead);
                            if (table != null)
                            {
                                for (int j = 0; j < table.Rows.Count; j += 16)
                                {
                                    for (int k = 0; k < 16; k++)
                                    {
                                        if (!table.Cells[j + k, 0].TextString.Contains("Резерв") && !table.Cells[j + k, 0].TextString.Contains("резерв") && table.Cells[j + k, 0].TextString != "")
                                        {
                                            UnitD unit = new UnitD();
                                            unit.Article = "6ES7322-1BL00-0AA0";
                                            unit.XT = "-XT" + currentRack + "." + currentSlot + "." + currentPart;
                                            unit.Equipment = table.Cells[j + k, 0].GetTextString(FormatOption.FormatOptionNone);
                                            unit.Equipment = unit.Equipment.Replace("\n", "");
                                            unit.Rack = currentRack;
                                            unit.Slot = currentSlot;
                                            unit.Type = "DO";
                                            unit.XR = "-" + currentPart + "XRI" + currentRack + "." + currentSlot;
                                            unit.LeftPin = currentPin;
                                            currentPin++;
                                            unit.RightPin = "OUT" + downPinPart + "." + downPinNum;
                                            downPinNum++;
                                            unit.XTPins = new List<string>();
                                            unit.XTLink = "-XT24.4:2.1/3.4";
                                            if (pages < fuCount)
                                            {
                                                unit.XTPins.Add("FU" + currentFCNum.ToString());
                                                unit.XTPins.Add(currentFCNum.ToString());
                                            }
                                            else
                                            {
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                                unit.XTPins.Add(currentXTPin.ToString());
                                                currentXTPin++;
                                            }
                                            unit.Channel = currentChannel;
                                            currentChannel++;

                                            unit.XRNum = currentFCNum;
                                            currentFCNum++;
                                            unit.XRs = new List<string>();
                                            unit.XRs.Add("A1");
                                            unit.XRs.Add("14");
                                            unit.XRs.Add("A2");
                                            unit.XRs.Add("11");

                                            unitsD.Add(unit);
                                        }
                                        else { currentPin++; downPinNum++; currentChannel++; currentXTPin += 3; currentFCNum++; }
                                        if (k == 7) { currentPin += 2; downPinNum = 0; downPinPart++; }
                                    }
                                    if (currentPart == 1) { currentPart++; currentPin += 2; downPinNum = 0; downPinPart = 3; currentFCNum = 1; currentXTPin = 1; }
                                    else
                                    {
                                        currentPart = 1;
                                        currentSlot++;
                                        currentPin = 2;
                                        currentModNum++;
                                        downPinPart = 1;
                                        downPinNum = 0;
                                        currentXTPin = 1;
                                        if (currentModNum == 10)
                                        {
                                            currentModNum = 1;
                                            currentRack++;
                                        }
                                        currentChannel = 0;
                                    }
                                    pages++;
                                    firstPage = false;
                                    currentFCNum = 1;
                                }
                            }
                        }
                    }
                }
            }

           
            List<string> data = new List<string>();
            for (int i = 0; i < unitsA.Count; i++)
            {
                string str = unitsA[i].Type + "░" + unitsA[i].Article + "░" + unitsA[i].Equipment + "░" + unitsA[i].Rack + "░" + unitsA[i].Slot + "░" + unitsA[i].Channel +"░" + unitsA[i].Terminals;
                str += "░{ ";
                if (unitsA[i].Fu!=null)
                    for (int j = 0; j < unitsA[i].Fu.Count; j++)
                    {
                        str += unitsA[i].Fu[j] + " ";
                    }
                str += "}░{";
                if (unitsA[i].Pins != null)
                    for (int j = 0; j < unitsA[i].Pins.Count; j++)
                    {
                        str += " " + unitsA[i].Pins[j];
                    }
                str += "}";
                data.Add(str);
            }

            for (int i = 0; i < unitsD.Count; i++)
            {
                string str = unitsD[i].Type + "░" + unitsD[i].Article + "░" + unitsD[i].Equipment + "░" + unitsD[i].Rack + "░" + unitsD[i].Slot + "░" + unitsD[i].Channel + "░" + unitsD[i].RightPin + "░" + unitsD[i].LeftPin + "░" + unitsD[i].XR + "░" + unitsD[i].XRNum;
                str += "░{ ";
                if (unitsD[i].XRs != null)
                    for (int j = 0; j < unitsD[i].XRs.Count; j++)
                    {
                        str += unitsD[i].XRs[j] + " ";
                    }
                str += "}░"+ unitsD[i].XT +"░{";
                if (unitsD[i].XTPins != null)
                    for (int j = 0; j < unitsD[i].XTPins.Count; j++)
                    {
                        str += unitsD[i].XTPins[j] + " ";
                    }
                str += "}░" + unitsD[i].XTLink;
                data.Add(str);
            }

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text|*.txt";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
                System.IO.File.WriteAllLines(saveFileDialog1.FileName, data);
        }

        private static void loadFonts(Transaction acTrans, BlockTableRecord modSpace, Database acdb)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            ObjectIdCollection id = new ObjectIdCollection();
            string filename = @"Data\ContourExample.dwg";
            using (Database sourceDb = new Database(false, true))
            {
                if (System.IO.File.Exists(filename))
                {
                    sourceDb.ReadDwgFile(filename, System.IO.FileShare.Read, true, "");
                    using (Transaction trans = sourceDb.TransactionManager.StartTransaction())
                    {
                        TextStyleTable table = (TextStyleTable)trans.GetObject(sourceDb.TextStyleTableId, OpenMode.ForRead);
                        if (table.Has("ROMANS0-60"))
                            id.Add(table["ROMANS0-60"]);
                        trans.Commit();
                    }
                }
                else editor.WriteMessage("Не найден файл {0}", filename);
                if (id.Count > 0)
                {
                    acTrans.TransactionManager.QueueForGraphicsFlush();
                    IdMapping iMap = new IdMapping();
                    acdb.WblockCloneObjects(id, acdb.TextStyleTableId, iMap, DuplicateRecordCloning.Ignore, false);
                }
            }
        }
    }

    public static class Transliteration
    {
        private static Dictionary<string, string> iso = new Dictionary<string, string>();

        public static string Front(string text)
        {
            string output = text;
            Dictionary<string, string> tdict = iso;

            foreach (KeyValuePair<string, string> key in tdict)
            {
                output = output.Replace(key.Key, key.Value);
            }
            return output;
        }

        static Transliteration()
        {
            iso.Add("Т", "T");
            iso.Add("Х", "X");
        }
    }
}
