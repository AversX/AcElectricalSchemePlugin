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

namespace AcElectricalSchemePlugin
{
    static class ConnectionScheme
    {
        private static Document acDoc;
        private static List<Unit> units;

        struct Unit
        {
            public string cupboardName;
            public string tBoxName;
            public string designation;
            public string param;
            public string equipment;
            public string equipType;
            public string cableMark;
            public string shield;
            public List<Terminal> terminals;
            public List<string> boxTerminals;
            public List<string> equipTerminals;

            public Unit(string _cupboardName, string _tBoxName, string _designation, string _param, string _equipment, string _equipType, string _cableMark, string _shield, List<Terminal> _terminals, List<string> _boxTerminals, List<string> _equipTerminals)
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
                boxTerminals = _boxTerminals;
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
            OpenFileDialog file = new OpenFileDialog();
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
                        List<string> boxTerminals = new List<string>();
                        for (int j = 18; j < 28; j++)
                            if (text[j] != "")
                                boxTerminals.Add(text[j]);
                        List<string> equipTerminals = new List<string>();
                        for (int j = 28; j < 38; j++)
                            if (text[j] != "")
                                equipTerminals.Add(text[j]);
                        units.Add(new Unit(text[0], text[1], text[2], text[3], text[4], text[5], text[6], text[7], terminals, boxTerminals, equipTerminals));
                    }
                    connectionScheme();
                }
            }
        }
        private static void connectionScheme()
        {
            Database acDb = acDoc.Database;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                acDoc.LockDocument();
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    #region shield
                    Polyline shieldPoly = new Polyline();
                    shieldPoly.SetDatabaseDefaults();
                    shieldPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    shieldPoly.Closed = true;
                    shieldPoly.AddVertexAt(0, new Point2d(1200, 2000), 0, 0, 0);
                    shieldPoly.AddVertexAt(1, new Point2d(1730, 2000), 0, 0, 0);
                    shieldPoly.AddVertexAt(2, new Point2d(1730, 1965), 0, 0, 0);
                    shieldPoly.AddVertexAt(3, new Point2d(1200, 1965), 0, 0, 0);
                    acModSpace.AppendEntity(shieldPoly);
                    acTrans.AddNewlyCreatedDBObject(shieldPoly, true);
                    #endregion
                    #region ground
                    Polyline gndPoly = new Polyline();
                    gndPoly.SetDatabaseDefaults();
                    gndPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    gndPoly.Closed = true;
                    gndPoly.AddVertexAt(0, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -21)), 0, 0, 0);
                    gndPoly.AddVertexAt(1, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(523, -21)), 0, 0, 0);
                    gndPoly.AddVertexAt(2, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(523, -23)), 0, 0, 0);
                    gndPoly.AddVertexAt(3, shieldPoly.GetPoint2dAt(0).Add(new Vector2d(5, -23)), 0, 0, 0);
                    acModSpace.AppendEntity(gndPoly);
                    acTrans.AddNewlyCreatedDBObject(gndPoly, true);
                    #endregion
                    #region TerminalBox
                    drawUnits(acTrans, acModSpace, units, gndPoly);
                    #endregion
                    #region Text
                    //DBText acText1 = new DBText();
                    //acText1.Position = new Point3d(101, 1070, 0);
                    //acText1.Height = 2.1;
                    //acText1.TextString = "ЩИТ";
                    //acText1.WidthFactor = 0.5;
                    //acModSpace.AppendEntity(acText1);
                    //acTrans.AddNewlyCreatedDBObject(acText1, true);

                    //DBText acText2 = new DBText();
                    //acText2.Position = new Point3d(101, 1067, 0);
                    //acText2.Height = 2.1;
                    //acText2.TextString = "СИЛОВОЙ";
                    //acText2.WidthFactor = 0.5;
                    //acModSpace.AppendEntity(acText2);
                    //acTrans.AddNewlyCreatedDBObject(acText2, true);

                    //DBText acText3 = new DBText();
                    //acText3.Position = new Point3d(101, 1064, 0);
                    //acText3.Height = 2.1;
                    //acText3.TextString = "ШС6";
                    //acText3.WidthFactor = 0.5;
                    //acModSpace.AppendEntity(acText3);
                    //acTrans.AddNewlyCreatedDBObject(acText3, true);
                    #endregion
 

                    acTrans.Commit();
                    acTrans.Dispose();
                }
            }
        }

        private static void drawUnits(Transaction acTrans, BlockTableRecord modSpace, List<Unit> units, Polyline gnd)
        {
            Polyline tBoxPoly = new Polyline();
            tBoxPoly.SetDatabaseDefaults();
            tBoxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tBoxPoly.Closed = true;
            tBoxPoly.AddVertexAt(0, gnd.GetPoint2dAt(3).Add(new Vector2d(0, -3)), 0, 0, 0);
            tBoxPoly.AddVertexAt(1, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(27, 0)), 0, 0, 0);
            tBoxPoly.AddVertexAt(2, tBoxPoly.GetPoint2dAt(1).Add(new Vector2d(0, -9)), 0, 0, 0);
            tBoxPoly.AddVertexAt(3, tBoxPoly.GetPoint2dAt(0).Add(new Vector2d(0, -9)), 0, 0, 0);
            modSpace.AppendEntity(tBoxPoly);
            acTrans.AddNewlyCreatedDBObject(tBoxPoly, true);
            Polyline prevPoly = tBoxPoly;
            Polyline prevTermPoly = null;
            string prevTerminal=null;
            for (int j = 0; j < units.Count; j++)
            {
                double leftEdgeX = 0;
                double rightEdgeX = 0;
                Point3d lowestPoint = Point3d.Origin;
                for (int i = 0; i < units[j].terminals.Count; i++)
                {
                    string terminalTag = units[j].terminals[i].boxTerminal1.Split('[', ']')[1];
                    string terminal1 = units[j].terminals[i].boxTerminal1.Split('[', ']')[2];
                    string terminal2 = units[j].terminals[i].boxTerminal2.Split('[', ']')[2];
                    if (prevTerminal == null)
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

                        prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal1, terminal2, false, out lowestPoint);
                        leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
                    }
                    else if (prevTerminal == terminalTag)
                    {
                        Point2d p1 = prevPoly.GetPoint2dAt(1).Add(new Vector2d(37, 0));
                        prevPoly.SetPointAt(1, p1);
                        Point2d p2 = prevPoly.GetPoint2dAt(2).Add(new Vector2d(37, 0));
                        prevPoly.SetPointAt(2, p2);
                        
                        prevTermPoly = drawTerminal(acTrans, modSpace, prevTermPoly, terminal1, terminal2, true, out lowestPoint);
                        if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
                    }
                    else
                    {
                        tBoxPoly = new Polyline();
                        tBoxPoly.SetDatabaseDefaults();
                        tBoxPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                        tBoxPoly.Closed = true;
                        tBoxPoly.AddVertexAt(0, prevPoly.GetPoint2dAt(1).Add(new Vector2d(15, 0)), 0, 0, 0);
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

                        prevTermPoly = drawTerminal(acTrans, modSpace, prevPoly, terminal1, terminal2, false, out lowestPoint);
                        if (i == 0) leftEdgeX = prevTermPoly.GetPoint2dAt(0).X - 6;
                    }
                }
                rightEdgeX = prevTermPoly.GetPoint2dAt(1).X;
                Point3d center = new Point3d(leftEdgeX + (rightEdgeX - leftEdgeX) / 2, prevTermPoly.GetPoint2dAt(2).Y-3, 0);
                Vector3d normal = Vector3d.ZAxis;
                Vector3d majorAxis = new Vector3d((rightEdgeX - leftEdgeX) / 2, 0, 0);
                double radiusRatio = 1.64/majorAxis.X;
                double startAngle = 0.0;
                double endAngle = Math.PI * 2;
                Ellipse groundEllipse = new Ellipse(center, normal, majorAxis, radiusRatio, startAngle, endAngle);
                modSpace.AppendEntity(groundEllipse);
                acTrans.AddNewlyCreatedDBObject(groundEllipse, true);

                Line groundLine1 = new Line();
                groundLine1.SetDatabaseDefaults();
                groundLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLine1.StartPoint = new Point3d(groundEllipse.Center.X+groundEllipse.MajorAxis.X, prevTermPoly.GetPoint2dAt(2).Y - 3, 0);
                groundLine1.EndPoint = groundLine1.StartPoint.Add(new Vector3d(6, 0, 0));
                modSpace.AppendEntity(groundLine1);
                acTrans.AddNewlyCreatedDBObject(groundLine1, true);

                Line groundLine2 = new Line();
                groundLine2.SetDatabaseDefaults();
                groundLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                groundLine2.StartPoint = groundLine1.EndPoint;
                groundLine2.EndPoint = groundLine2.StartPoint.Add(new Vector3d(0, gnd.StartPoint.Y - groundLine2.StartPoint.Y - 1.5, 0));
                modSpace.AppendEntity(groundLine2);
                acTrans.AddNewlyCreatedDBObject(groundLine2, true);

                Circle groundCircle = new Circle();
                groundCircle.SetDatabaseDefaults();
                groundCircle.Center = new Point3d(groundLine2.EndPoint.X, groundLine2.EndPoint.Y+0.36, 0);
                groundCircle.Radius = 0.36;
                modSpace.AppendEntity(groundCircle);
                acTrans.AddNewlyCreatedDBObject(groundCircle, true);

                drawCable(acTrans, modSpace, new Point3d(leftEdgeX+3, lowestPoint.Y, 0), new Point3d(rightEdgeX-3, lowestPoint.Y, 0), units[j]);
            }
        }
        private static Polyline drawTerminal(Transaction acTrans, BlockTableRecord modSpace, Polyline prevPoly, string term1, string term2, bool iteration, out Point3d lowestPoint)
        {
            Polyline termPoly1 = new Polyline();
            termPoly1.SetDatabaseDefaults();
            termPoly1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            termPoly1.Closed = true;
            termPoly1.AddVertexAt(0, prevPoly.GetPoint2dAt(0).Add(new Vector2d(iteration?31:15, 0)), 0, 0, 0);
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

            Line cableLineUp2 = new Line();
            cableLineUp2.SetDatabaseDefaults();
            cableLineUp2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            cableLineUp2.StartPoint = new Point3d(termPoly2.GetPoint2dAt(3).X + (termPoly2.GetPoint2dAt(2).X - termPoly2.GetPoint2dAt(3).X) / 2, termPoly2.GetPoint2dAt(3).Y, 0);
            cableLineUp2.EndPoint = cableLineUp2.StartPoint.Add(new Vector3d(0, -44, 0));
            modSpace.AppendEntity(cableLineUp2);
            acTrans.AddNewlyCreatedDBObject(cableLineUp2, true);

            lowestPoint = cableLineUp2.EndPoint;

            return termPoly2;
        }

        private static void drawCable(Transaction acTrans, BlockTableRecord modSpace, Point3d firstCable, Point3d lastCable, Unit unit)
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

            
            Line cableLine2 = new Line();
            cableLine2.SetDatabaseDefaults();
            cableLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            startPoint = new Point3d(endPoint.X, endPoint.Y - 10, 0);
            double cableLine2Length;
            try
            {
                cableLine2Length = unit.boxTerminals.Count > 0 ? startPoint.Y - 40 : startPoint.Y - 80;
            }
            catch
            {
                cableLine2Length = startPoint.Y - 80;
            }
            endPoint = new Point3d(startPoint.X, cableLine2Length, 0);
            cableLine2.StartPoint = startPoint;
            cableLine2.EndPoint = endPoint;
            modSpace.AppendEntity(cableLine2);
            acTrans.AddNewlyCreatedDBObject(cableLine2, true);

            Point3d point;
            if (unit.boxTerminals != null && unit.boxTerminals.Count > 0)
            {
                point = drawTerminalBox(acTrans, modSpace, cableLine2, unit);
            }
            else
            {
                point = cableLine2.EndPoint;
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

                DBText text = new DBText();
                text.SetDatabaseDefaults();
                text.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                text.Position = terminal.GetPoint3dAt(0).Add(new Vector3d(2.5, -3.5, 0));
                text.TextString = unit.equipTerminals[i];
                text.HorizontalMode = TextHorizontalMode.TextCenter;
                text.AlignmentPoint = text.Position;
                modSpace.AppendEntity(text);
                acTrans.AddNewlyCreatedDBObject(text, true);
            }

            Polyline equip = new Polyline();
            equip.SetDatabaseDefaults();
            equip.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            equip.Closed = true;
            equip.AddVertexAt(0, new Point2d(jumperLineDown.StartPoint.X - 10, jumperLineDown.StartPoint.Y + 3.5), 0, 0, 0);
            equip.AddVertexAt(1, new Point2d(jumperLineDown.EndPoint.X + 10, jumperLineDown.StartPoint.Y + 3.5), 0, 0, 0);
            equip.AddVertexAt(2, equip.GetPoint2dAt(1).Add(new Vector2d(0, -60)), 0, 0, 0);
            equip.AddVertexAt(3, equip.GetPoint2dAt(0).Add(new Vector2d(0, -60)), 0, 0, 0);
            modSpace.AppendEntity(equip);
            acTrans.AddNewlyCreatedDBObject(equip, true);
        }

        private static Point3d drawTerminalBox(Transaction acTrans, BlockTableRecord modSpace, Line cableLine, Unit unit)
        {
            Line jumperLineDown = new Line();
            jumperLineDown.SetDatabaseDefaults();
            jumperLineDown.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            double x = cableLine.EndPoint.X - ((unit.boxTerminals.Count-1) * 5) / 2;
            jumperLineDown.StartPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            x = x + ((unit.boxTerminals.Count-1) * 5);
            jumperLineDown.EndPoint = new Point3d(x, cableLine.EndPoint.Y, 0);
            modSpace.AppendEntity(jumperLineDown);
            acTrans.AddNewlyCreatedDBObject(jumperLineDown, true);

            double lowestPointY = 0;
            double leftEdgeX = 0;
            double rightEdgeX = 0;
            for (int i = 0; i < unit.boxTerminals.Count; i++)
            {
                if (unit.boxTerminals[i] != "*")
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
                    text.TextString = unit.boxTerminals[i];
                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                    text.AlignmentPoint = text.Position;
                    modSpace.AppendEntity(text);
                    acTrans.AddNewlyCreatedDBObject(text, true);

                    Line cableLineOutput = new Line();
                    cableLineOutput.SetDatabaseDefaults();
                    cableLineOutput.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    cableLineOutput.StartPoint = cableLineInput.EndPoint.Add(new Vector3d(0, -34, 0));
                    cableLineOutput.EndPoint = cableLineOutput.StartPoint.Add(new Vector3d(0, -8, 0));
                    modSpace.AppendEntity(cableLineOutput);
                    acTrans.AddNewlyCreatedDBObject(cableLineOutput, true);

                    lowestPointY = cableLineOutput.EndPoint.Y;
                    if (i == 0) leftEdgeX = cableLineInput.StartPoint.X - 3;
                    rightEdgeX = cableLineInput.StartPoint.X + 3;
                }
                else
                {
                    Polyline terminal = new Polyline();
                    terminal.SetDatabaseDefaults();
                    terminal.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    terminal.Closed = true;
                    terminal.AddVertexAt(0, new Point2d(rightEdgeX - 0.75, jumperLineDown.EndPoint.Y - 8), 0, 0, 0);
                    terminal.AddVertexAt(1, terminal.GetPoint2dAt(0).Add(new Vector2d(5, 0)), 0, 0, 0);
                    terminal.AddVertexAt(2, terminal.GetPoint2dAt(1).Add(new Vector2d(0, -34)), 0, 0, 0);
                    terminal.AddVertexAt(3, terminal.GetPoint2dAt(0).Add(new Vector2d(0, -34)), 0, 0, 0);
                    modSpace.AppendEntity(terminal);
                    acTrans.AddNewlyCreatedDBObject(terminal, true);

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
                    text.TextString = unit.boxTerminals[i];
                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                    text.AlignmentPoint = text.Position;
                    modSpace.AppendEntity(text);
                    acTrans.AddNewlyCreatedDBObject(text, true);

                    Point3d center = new Point3d(leftEdgeX + (rightEdgeX - leftEdgeX) / 2, jumperLineDown.StartPoint.Y - 4, 0);
                    Vector3d normal = Vector3d.ZAxis;
                    Vector3d majorAxis = new Vector3d((rightEdgeX - leftEdgeX) / 2, 0, 0);
                    double radiusRatio = 1.64 / majorAxis.X;
                    double startAngle = 0.0;
                    double endAngle = Math.PI * 2;
                    Ellipse groundEllipse = new Ellipse(center, normal, majorAxis, radiusRatio, startAngle, endAngle);
                    modSpace.AppendEntity(groundEllipse);
                    acTrans.AddNewlyCreatedDBObject(groundEllipse, true);

                    Line groundLine1 = new Line();
                    groundLine1.SetDatabaseDefaults();
                    groundLine1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    groundLine1.StartPoint = new Point3d(groundEllipse.Center.X + groundEllipse.MajorAxis.X, jumperLineDown.StartPoint.Y - 4, 0);
                    groundLine1.EndPoint = groundLine1.StartPoint.Add(new Vector3d(2.5, 0, 0));
                    modSpace.AppendEntity(groundLine1);
                    acTrans.AddNewlyCreatedDBObject(groundLine1, true);

                    Line groundLine2 = new Line();
                    groundLine2.SetDatabaseDefaults();
                    groundLine2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    groundLine2.StartPoint = groundLine1.EndPoint;
                    groundLine2.EndPoint = groundLine2.StartPoint.Add(new Vector3d(0, terminal.GetPoint2dAt(0).Y - groundLine2.StartPoint.Y, 0));
                    modSpace.AppendEntity(groundLine2);
                    acTrans.AddNewlyCreatedDBObject(groundLine2, true);

                    jumperLineDown.EndPoint = jumperLineDown.EndPoint.Add(new Vector3d(-5, 0, 0));

                    center = new Point3d(leftEdgeX + (rightEdgeX - leftEdgeX) / 2, lowestPointY + 4, 0);
                    normal = Vector3d.ZAxis;
                    majorAxis = new Vector3d((rightEdgeX - leftEdgeX) / 2, 0, 0);
                    radiusRatio = 1.64 / majorAxis.X;
                    startAngle = 0.0;
                    endAngle = Math.PI * 2;
                    Ellipse groundEllipseDown = new Ellipse(center, normal, majorAxis, radiusRatio, startAngle, endAngle);
                    modSpace.AppendEntity(groundEllipseDown);
                    acTrans.AddNewlyCreatedDBObject(groundEllipseDown, true);

                    Line groundLineDown1 = new Line();
                    groundLineDown1.SetDatabaseDefaults();
                    groundLineDown1.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    groundLineDown1.StartPoint = new Point3d(groundEllipseDown.Center.X + groundEllipseDown.MajorAxis.X, lowestPointY + 4, 0);
                    groundLineDown1.EndPoint = groundLineDown1.StartPoint.Add(new Vector3d(2.5, 0, 0));
                    modSpace.AppendEntity(groundLineDown1);
                    acTrans.AddNewlyCreatedDBObject(groundLineDown1, true);

                    Line groundLineDown2 = new Line();
                    groundLineDown2.SetDatabaseDefaults();
                    groundLineDown2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    groundLineDown2.StartPoint = groundLineDown1.EndPoint;
                    groundLineDown2.EndPoint = groundLineDown2.StartPoint.Add(new Vector3d(0, groundLine2.StartPoint.Y - terminal.GetPoint2dAt(0).Y, 0));
                    modSpace.AppendEntity(groundLineDown2);
                    acTrans.AddNewlyCreatedDBObject(groundLineDown2, true);
                }
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

            Polyline tbox = new Polyline();
            tbox.SetDatabaseDefaults();
            tbox.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            tbox.Closed = true;
            tbox.AddVertexAt(0, new Point2d(jumperLineDown.StartPoint.X - 7, jumperLineDown.StartPoint.Y + 2), 0, 0, 0);
            tbox.AddVertexAt(1, new Point2d(jumperLineDown.EndPoint.X + 9, jumperLineDown.StartPoint.Y + 2), 0, 0, 0);
            tbox.AddVertexAt(2, tbox.GetPoint2dAt(1).Add(new Vector2d(0, -53.35)), 0, 0, 0);
            tbox.AddVertexAt(3, tbox.GetPoint2dAt(0).Add(new Vector2d(0, -53.35)), 0, 0, 0);
            modSpace.AppendEntity(tbox);
            acTrans.AddNewlyCreatedDBObject(tbox, true);
            return cLineOutput.EndPoint;
        }
    }
}
