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
        private static Unit[] units;

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
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            //OpenFileDialog file = new OpenFileDialog();
            //if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    string[] lines = System.IO.File.ReadAllLines(file.FileName);
            //    if (lines.Length > 0)
            //    //    string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + file.FileName + "; Extended Properties=Excel 8.0;";
            //    //    OleDbConnection connection = new OleDbConnection(connectionString);
            //    //    connection.Open();
            //    //    string sheetName = (string)connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }).Rows[0].ItemArray[2];
            //    //    OleDbCommand select = new OleDbCommand("SELECT * FROM [" + sheetName + "]", connection);
            //    //    OleDbDataAdapter adapter = new OleDbDataAdapter();
            //    //    adapter.SelectCommand = select;
            //    //    DataSet dSet = new DataSet();
            //    //    adapter.Fill(dSet);
            //    //    if (dSet.Tables[0].Rows.Count > 0)
            //    {
            //        //        Unit[] units = new Unit[dSet.Tables[0].Rows.Count];
            //        //        for (int i = 0; i < dSet.Tables[0].Rows.Count; i++)
            //        //        {
            //        //            units[i].Name = dSet.Tables[0].Rows[i][0].ToString();
            //        //            units[i].Terminals = int.Parse(dSet.Tables[0].Rows[i][1].ToString());
            //        //        }
            //        units = new Unit[lines.Length];
            //        for (int i = 0; i < lines.Length; i++)
            //        {
            //            string[] text = lines[i].Split(' ');
            //            units[i].Name = text[0].ToString();
            //            units[i].Terminals = int.Parse(text[1].ToString());
            //        }
            //        greenLine();
                    connectionScheme();
            //    }

            //}
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
                    List<Unit> units = new List<Unit>();
                    
                    Unit unit = new Unit();
                    List<Terminal> terminals = new List<Terminal>();
                    terminals.Add(new Terminal("[XT3.6]FU1", "[XT3.6]1"));
                    unit.equipTerminals = new List<string>{"+","-"};
                    unit.terminals = terminals;
                    units.Add(unit);

                    unit = new Unit();
                    terminals = new List<Terminal>();
                    terminals.Add(new Terminal("[XT3.6]FU2", "[XT3.6]2"));
                    unit.equipTerminals = new List<string> { "+", "-" };
                    unit.terminals = terminals;
                    units.Add(unit);

                    unit = new Unit();
                    terminals = new List<Terminal>();
                    terminals.Add(new Terminal("[XT3.6]FU3", "[XT3.6]3"));
                    unit.equipTerminals = new List<string> { "+", "-" };
                    unit.boxTerminals = new List<string> { "1", "2", "1", "2", "1", "2", "*" };
                    unit.terminals = terminals;
                    units.Add(unit);

                    unit = new Unit();
                    terminals = new List<Terminal>();
                    terminals.Add(new Terminal("[XT24.1]FU1", "[XT24.1]1"));
                    terminals.Add(new Terminal("[XT4.4]FU1", "[XT4.4]1"));
                    terminals.Add(new Terminal("[XT3.6]FU5", "[XT3.6]5"));
                    unit.equipTerminals = new List<string> { "+", "-", "+", "-" };
                    unit.boxTerminals = new List<string> { "1", "2", "*" };
                    unit.terminals = terminals;
                    units.Add(unit);

                    unit = new Unit();
                    terminals = new List<Terminal>();
                    terminals.Add(new Terminal("[XT4.4]FU2", "[XT4.4]2"));
                    terminals.Add(new Terminal("[XT3.6]FU6", "[XT3.6]6"));
                    unit.equipTerminals = new List<string> { "+", "-", "+", "-", "+", "-" };
                    unit.boxTerminals = new List<string> { "1", "2", "*" };
                    unit.terminals = terminals;
                    units.Add(unit);

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
                    #region connectors
                    //int terminals = units[0].Terminals;
                    //#region power
                    //double offset;
                    //for (int i = 0; i < terminals; i++)
                    //{
                    //    Polyline pinConnectorIn = new Polyline();
                    //    pinConnectorIn.SetDatabaseDefaults();
                    //    pinConnectorIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    pinConnectorIn.AddVertexAt(0, new Point2d(132 + 3.5 * i, 1060), 0, 0, 0);
                    //    pinConnectorIn.AddVertexAt(1, new Point2d(132 + 3.5 * i, 1065), 0, 0, 0);
                    //    pinConnectorIn.AddVertexAt(2, new Point2d(132 + 3.5 * i + 3.5, 1065), 0, 0, 0);
                    //    pinConnectorIn.AddVertexAt(3, new Point2d(132 + 3.5 * i + 3.5, 1060), 0, 0, 0);
                    //    pinConnectorIn.AddVertexAt(4, new Point2d(132 + 3.5 * i, 1060), 0, 0, 0);
                    //    acModSpace.AppendEntity(pinConnectorIn);
                    //    acTrans.AddNewlyCreatedDBObject(pinConnectorIn, true);

                    //    Polyline powerConnectorIn = new Polyline();
                    //    powerConnectorIn.SetDatabaseDefaults();
                    //    powerConnectorIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    powerConnectorIn.AddVertexAt(0, new Point2d(133.75 + 3.5 * i, 1060), 0, 0, 0);
                    //    powerConnectorIn.AddVertexAt(1, new Point2d(133.75 + 3.5 * i, 1055), 0, 0, 0);
                    //    acModSpace.AppendEntity(powerConnectorIn);
                    //    acTrans.AddNewlyCreatedDBObject(powerConnectorIn, true);

                    //    Polyline powerConnectorOut = new Polyline();
                    //    powerConnectorOut.SetDatabaseDefaults();
                    //    powerConnectorOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    powerConnectorOut.AddVertexAt(0, new Point2d(133.75 + 3.5 * i, 985), 0, 0, 0);
                    //    powerConnectorOut.AddVertexAt(1, new Point2d(133.75 + 3.5 * i, 980), 0, 0, 0);
                    //    acModSpace.AppendEntity(powerConnectorOut);
                    //    acTrans.AddNewlyCreatedDBObject(powerConnectorOut, true);

                    //    Circle powerCircle = new Circle();
                    //    powerCircle.SetDatabaseDefaults();
                    //    powerCircle.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    powerCircle.Center = new Point3d(133.75 + 3.5 * i, 979.3, 0);
                    //    powerCircle.Radius = 0.7;
                    //    acModSpace.AppendEntity(powerCircle);
                    //    acTrans.AddNewlyCreatedDBObject(powerCircle, true);

                    //    Line powerCircleLine = new Line(new Point3d(132.75 + 3.5 * i, 978.3, 0), new Point3d(134.75 + 3.5 * i, 980.3, 0));
                    //    powerCircleLine.SetDatabaseDefaults();
                    //    powerCircleLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    acModSpace.AppendEntity(powerCircleLine);
                    //    acTrans.AddNewlyCreatedDBObject(powerCircleLine, true);

                    //    Polyline circleToPin = new Polyline();
                    //    circleToPin.SetDatabaseDefaults();
                    //    circleToPin.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    circleToPin.AddVertexAt(0, new Point2d(133.75 + 3.5 * i, 978.6), 0, 0, 0);
                    //    circleToPin.AddVertexAt(1, new Point2d(133.75 + 3.5 * i, 976), 0, 0, 0);
                    //    acModSpace.AppendEntity(circleToPin);
                    //    acTrans.AddNewlyCreatedDBObject(circleToPin, true);

                    //    Polyline pinConnectorOut = new Polyline();
                    //    pinConnectorOut.SetDatabaseDefaults();
                    //    pinConnectorOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    pinConnectorOut.AddVertexAt(0, new Point2d(132 + 3.5 * i, 976), 0, 0, 0);
                    //    pinConnectorOut.AddVertexAt(1, new Point2d(132 + 3.5 * i + 3.5, 976), 0, 0, 0);
                    //    pinConnectorOut.AddVertexAt(2, new Point2d(132 + 3.5 * i + 3.5, 971), 0, 0, 0);
                    //    pinConnectorOut.AddVertexAt(3, new Point2d(132 + 3.5 * i, 971), 0, 0, 0);
                    //    pinConnectorOut.AddVertexAt(4, new Point2d(132 + 3.5 * i, 976), 0, 0, 0);
                    //    acModSpace.AppendEntity(pinConnectorOut);
                    //    acTrans.AddNewlyCreatedDBObject(pinConnectorOut, true);
                    //}
                    //#region powerConnectorAllIn
                    //Polyline powerConnectorAllIn = new Polyline();
                    //powerConnectorAllIn.SetDatabaseDefaults();
                    //powerConnectorAllIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //powerConnectorAllIn.AddVertexAt(0, new Point2d(133.75, 1055), 0, 0, 0);
                    //powerConnectorAllIn.AddVertexAt(1, new Point2d(133.75 + 3.5 * (terminals - 1), 1055), 0, 0, 0);
                    //acModSpace.AppendEntity(powerConnectorAllIn);
                    //acTrans.AddNewlyCreatedDBObject(powerConnectorAllIn, true);
                    //#endregion
                    //#region powerArcIn
                    //Arc powerArcIn = new Arc(new Point3d(133.75 + 3.5 * (terminals - 1) / 2, 1055, 0), 1.75, 3.14, 0);
                    //powerArcIn.SetDatabaseDefaults();
                    //acModSpace.AppendEntity(powerArcIn);
                    //acTrans.AddNewlyCreatedDBObject(powerArcIn, true);
                    //#endregion
                    //#region powerConnectorAllOut
                    //Polyline powerConnectorAllOut = new Polyline();
                    //powerConnectorAllOut.SetDatabaseDefaults();
                    //powerConnectorAllOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //powerConnectorAllOut.AddVertexAt(0, new Point2d(133.75, 985), 0, 0, 0);
                    //powerConnectorAllOut.AddVertexAt(1, new Point2d(133.75 + 3.5 * (terminals - 1), 985), 0, 0, 0);
                    //acModSpace.AppendEntity(powerConnectorAllOut);
                    //acTrans.AddNewlyCreatedDBObject(powerConnectorAllOut, true);
                    //#endregion
                    //#region powerArcOut
                    //Arc powerArcOut = new Arc(new Point3d(133.75 + 3.5 * (terminals - 1) / 2, 985, 0), 1.75, 0, 3.14);
                    //powerArcOut.SetDatabaseDefaults();
                    //acModSpace.AppendEntity(powerArcOut);
                    //acTrans.AddNewlyCreatedDBObject(powerArcOut, true);
                    //#endregion
                    ////штриховкой    переделать
                    //#region powerAllPinsPoly
                    //Polyline powerAllPinsPoly = new Polyline();
                    //powerAllPinsPoly.SetDatabaseDefaults();
                    //powerAllPinsPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //powerAllPinsPoly.AddVertexAt(0, new Point2d(130, 981), 0, 0, 0);
                    //powerAllPinsPoly.AddVertexAt(1, new Point2d(130 + 3.5 * (terminals - 1) + 8, 981), 0, 0, 0);
                    //powerAllPinsPoly.AddVertexAt(2, new Point2d(130 + 3.5 * (terminals - 1) + 8, 970), 0, 0, 0);
                    //powerAllPinsPoly.AddVertexAt(3, new Point2d(130, 970), 0, 0, 0);
                    //powerAllPinsPoly.AddVertexAt(4, new Point2d(130, 981), 0, 0, 0);
                    //acModSpace.AppendEntity(powerAllPinsPoly);
                    //acTrans.AddNewlyCreatedDBObject(powerAllPinsPoly, true);
                    //#endregion
                    //#region powerCablePoly
                    //Polyline powerCablePoly = new Polyline();
                    //powerCablePoly.SetDatabaseDefaults();
                    //powerCablePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //powerCablePoly.AddVertexAt(0, new Point2d(133.75 + 3.5 * (terminals - 1) / 2, 1053.25), 0, 0, 0);
                    //powerCablePoly.AddVertexAt(1, new Point2d(133.75 + 3.5 * (terminals - 1) / 2, 986.75), 0, 0, 0);
                    //acModSpace.AppendEntity(powerCablePoly);
                    //acTrans.AddNewlyCreatedDBObject(powerCablePoly, true);
                    //#endregion
                    //offset = 130 + 3.5 * (terminals - 1) + 8 + 20;
                    //#endregion
                    //#region equipment
                    //for (int i = 0; i < units.Length; i++)
                    //{
                    //    for (int j = 0; j < units[i].Terminals; j++)
                    //    {
                    //        Polyline pinConnectorIn = new Polyline();
                    //        pinConnectorIn.SetDatabaseDefaults();
                    //        pinConnectorIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        pinConnectorIn.AddVertexAt(0, new Point2d(offset + 3.5 * j, 1060), 0, 0, 0);
                    //        pinConnectorIn.AddVertexAt(1, new Point2d(offset + 3.5 * j, 1065), 0, 0, 0);
                    //        pinConnectorIn.AddVertexAt(2, new Point2d(offset + 3.5 * j + 3.5, 1065), 0, 0, 0);
                    //        pinConnectorIn.AddVertexAt(3, new Point2d(offset + 3.5 * j + 3.5, 1060), 0, 0, 0);
                    //        pinConnectorIn.AddVertexAt(4, new Point2d(offset + 3.5 * j, 1060), 0, 0, 0);
                    //        acModSpace.AppendEntity(pinConnectorIn);
                    //        acTrans.AddNewlyCreatedDBObject(pinConnectorIn, true);

                    //        Polyline connectorIn = new Polyline();
                    //        connectorIn.SetDatabaseDefaults();
                    //        connectorIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        connectorIn.AddVertexAt(0, new Point2d(offset + 1.75 + 3.5 * j, 1060), 0, 0, 0);
                    //        connectorIn.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * j, 1055), 0, 0, 0);
                    //        acModSpace.AppendEntity(connectorIn);
                    //        acTrans.AddNewlyCreatedDBObject(connectorIn, true);

                    //        Polyline connectorOut = new Polyline();
                    //        connectorOut.SetDatabaseDefaults();
                    //        connectorOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        connectorOut.AddVertexAt(0, new Point2d(offset + 1.75 + 3.5 * j, 985), 0, 0, 0);
                    //        connectorOut.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * j, 980), 0, 0, 0);
                    //        acModSpace.AppendEntity(connectorOut);
                    //        acTrans.AddNewlyCreatedDBObject(connectorOut, true);

                    //        Circle circlePin = new Circle();
                    //        circlePin.SetDatabaseDefaults();
                    //        circlePin.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        circlePin.Center = new Point3d(offset + 1.75 + 3.5 * j, 979.3, 0);
                    //        circlePin.Radius = 0.7;
                    //        acModSpace.AppendEntity(circlePin);
                    //        acTrans.AddNewlyCreatedDBObject(circlePin, true);

                    //        Line circlePinLine = new Line(new Point3d(offset + 0.75 + 3.5 * j, 978.3, 0), new Point3d(offset + 2.75 + 3.5 * j, 980.3, 0));
                    //        circlePinLine.SetDatabaseDefaults();
                    //        circlePinLine.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        acModSpace.AppendEntity(circlePinLine);
                    //        acTrans.AddNewlyCreatedDBObject(circlePinLine, true);

                    //        Polyline circleToPin = new Polyline();
                    //        circleToPin.SetDatabaseDefaults();
                    //        circleToPin.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        circleToPin.AddVertexAt(0, new Point2d(offset + 1.75 + 3.5 * j, 978.6), 0, 0, 0);
                    //        circleToPin.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * j, 976), 0, 0, 0);
                    //        acModSpace.AppendEntity(circleToPin);
                    //        acTrans.AddNewlyCreatedDBObject(circleToPin, true);

                    //        Polyline pinConnectorOut = new Polyline();
                    //        pinConnectorOut.SetDatabaseDefaults();
                    //        pinConnectorOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //        pinConnectorOut.AddVertexAt(0, new Point2d(offset + 3.5 * j, 976), 0, 0, 0);
                    //        pinConnectorOut.AddVertexAt(1, new Point2d(offset + 3.5 * j + 3.5, 976), 0, 0, 0);
                    //        pinConnectorOut.AddVertexAt(2, new Point2d(offset + 3.5 * j + 3.5, 971), 0, 0, 0);
                    //        pinConnectorOut.AddVertexAt(3, new Point2d(offset + 3.5 * j, 971), 0, 0, 0);
                    //        pinConnectorOut.AddVertexAt(4, new Point2d(offset + 3.5 * j, 976), 0, 0, 0);
                    //        acModSpace.AppendEntity(pinConnectorOut);
                    //        acTrans.AddNewlyCreatedDBObject(pinConnectorOut, true);
                    //    }
                    //    #region connectorAllIn
                    //    Polyline connectorAllIn = new Polyline();
                    //    connectorAllIn.SetDatabaseDefaults();
                    //    connectorAllIn.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    connectorAllIn.AddVertexAt(0, new Point2d(offset + 1.75, 1055), 0, 0, 0);
                    //    connectorAllIn.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * (units[i].Terminals - 1), 1055), 0, 0, 0);
                    //    acModSpace.AppendEntity(connectorAllIn);
                    //    acTrans.AddNewlyCreatedDBObject(connectorAllIn, true);
                    //    #endregion
                    //    #region arcIn
                    //    Arc arcIn = new Arc(new Point3d(offset + 1.75 + 3.5 * (units[i].Terminals - 1) / 2, 1055, 0), 1.75, 3.14, 0);
                    //    arcIn.SetDatabaseDefaults();
                    //    acModSpace.AppendEntity(arcIn);
                    //    acTrans.AddNewlyCreatedDBObject(arcIn, true);
                    //    #endregion
                    //    #region connectorAllOut
                    //    Polyline connectorAllOut = new Polyline();
                    //    connectorAllOut.SetDatabaseDefaults();
                    //    connectorAllOut.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    connectorAllOut.AddVertexAt(0, new Point2d(offset + 1.75, 985), 0, 0, 0);
                    //    connectorAllOut.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * (units[i].Terminals - 1), 985), 0, 0, 0);
                    //    acModSpace.AppendEntity(connectorAllOut);
                    //    acTrans.AddNewlyCreatedDBObject(connectorAllOut, true);
                    //    #endregion
                    //    #region powerArcOut
                    //    Arc arcOut = new Arc(new Point3d(offset + 1.75 + 3.5 * (units[i].Terminals - 1) / 2, 985, 0), 1.75, 0, 3.14);
                    //    arcOut.SetDatabaseDefaults();
                    //    acModSpace.AppendEntity(arcOut);
                    //    acTrans.AddNewlyCreatedDBObject(arcOut, true);
                    //    #endregion
                    //    //штриховкой    переделать
                    //    #region allPinsPoly
                    //    Polyline allPinsPoly = new Polyline();
                    //    allPinsPoly.SetDatabaseDefaults();
                    //    allPinsPoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    allPinsPoly.AddVertexAt(0, new Point2d(offset - 2, 981), 0, 0, 0);
                    //    allPinsPoly.AddVertexAt(1, new Point2d(offset - 2 + 3.5 * (units[i].Terminals - 1) + 8, 981), 0, 0, 0);
                    //    allPinsPoly.AddVertexAt(2, new Point2d(offset - 2 + 3.5 * (units[i].Terminals - 1) + 8, 970), 0, 0, 0);
                    //    allPinsPoly.AddVertexAt(3, new Point2d(offset - 2, 970), 0, 0, 0);
                    //    allPinsPoly.AddVertexAt(4, new Point2d(offset - 2, 981), 0, 0, 0);
                    //    acModSpace.AppendEntity(allPinsPoly);
                    //    acTrans.AddNewlyCreatedDBObject(allPinsPoly, true);
                    //    #endregion
                    //    #region cablePoly
                    //    Polyline cablePoly = new Polyline();
                    //    cablePoly.SetDatabaseDefaults();
                    //    cablePoly.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
                    //    cablePoly.AddVertexAt(0, new Point2d(offset + 1.75 + 3.5 * (units[i].Terminals - 1) / 2, 1053.25), 0, 0, 0);
                    //    cablePoly.AddVertexAt(1, new Point2d(offset + 1.75 + 3.5 * (units[i].Terminals - 1) / 2, 986.75), 0, 0, 0);
                    //    acModSpace.AppendEntity(cablePoly);
                    //    acTrans.AddNewlyCreatedDBObject(cablePoly, true);
                    //    #endregion
                    //    offset = offset - 2 + 3.5 * (units[i].Terminals - 1) + 8 + 5;
                    //}
                    //#endregion
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
                    text.Position = terminal.GetPoint3dAt(0).Add(new Vector3d(2.5, -3.5, 0));
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
                    text.Position = terminal.GetPoint3dAt(0).Add(new Vector3d(2.5, -3.5, 0));
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

            //Line jumperLineDown2 = new Line();
            //jumperLineDown2.SetDatabaseDefaults();
            //jumperLineDown2.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            //x = cLineOutput.EndPoint.X - ((unit.equipTerminals.Count - 1) * 5) / 2;
            //jumperLineDown2.StartPoint = new Point3d(x, cLineOutput.EndPoint.Y, 0);
            //x = x + ((unit.equipTerminals.Count - 1) * 5);
            //jumperLineDown2.EndPoint = new Point3d(x, cLineOutput.EndPoint.Y, 0);
            //modSpace.AppendEntity(jumperLineDown2);
            //acTrans.AddNewlyCreatedDBObject(jumperLineDown2, true);

            //Polyline tbox = new Polyline();
            //tbox.SetDatabaseDefaults();
            //tbox.Color = Color.FromColorIndex(ColorMethod.ByLayer, 9);
            //tbox.Closed = true;
            //tbox.AddVertexAt(0, new Point2d(jumperLineDown2.StartPoint.X - 7, jumperLineDown2.StartPoint.Y + 12), 0, 0, 0);
            //tbox.AddVertexAt(1, new Point2d(jumperLineDown2.EndPoint.X + 3.5, jumperLineDown2.StartPoint.Y + 12), 0, 0, 0);
            //tbox.AddVertexAt(2, tbox.GetPoint2dAt(1).Add(new Vector2d(0, -43.35)), 0, 0, 0);
            //tbox.AddVertexAt(3, tbox.GetPoint2dAt(0).Add(new Vector2d(0, -43.35)), 0, 0, 0);
            //modSpace.AppendEntity(tbox);
            //acTrans.AddNewlyCreatedDBObject(tbox, true);
            return cLineOutput.EndPoint;
        }
    }
}
