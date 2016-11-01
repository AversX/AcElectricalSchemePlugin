using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
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
            public string Name;
            public int Terminals;
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
                    //#region Text
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
                    //#endregion
                    //#region connectors
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
                    //#endregion

                    acTrans.Commit();
                    acTrans.Dispose();
                }
            }
        }
    }
}
