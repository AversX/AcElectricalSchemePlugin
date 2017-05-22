using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;
using System.Data;
using System.Data.OleDb;

namespace AcElectricalSchemePlugin
{
    static class TerminalsWeightClass
    {
        private static Editor editor;
        private static Document acDoc;
        private static Database acDb;
        private static List<Unit> Units;
        private static List<Unit> Weights;

        struct Unit
        {
            public string Terminal;
            public string Name;
            public string Weight;
            public Unit(string terminal, string weight, string name)
            {
                Terminal = terminal;
                Weight = weight;
                Name = name;
            }
        }

        static public void Calculate()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            acDb = acDoc.Database;

            Weights = LoadWeights();

            TypedValue[] filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "INSERT");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ТЕКСТ");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = acDoc.Editor.GetSelection(filter);

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    Units = getAllUnits(acTrans, selRes);
                    for (int i=0; i<Units.Count; i++)
                    {
                        bool passed = false;
                        for (int j = 0; j < Weights.Count; j++)
                            if (Units[i].Name == Weights[j].Name)
                            {
                                passed = true;
                                Unit unit = Units[i];
                                unit.Weight = Weights[j].Weight;
                                Units[i] = unit;
                                break;
                            }
                        if (!passed) editor.WriteMessage("\n\nНе найден вес для клеммы " + Units[i].Terminal + " Наименование " + Units[i].Name);
                    }
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text|*.txt";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
                System.IO.File.WriteAllLines(saveFileDialog1.FileName, getData());
        }

        private static List<Unit> LoadWeights()
        {
            List<Unit> weights = new List<Unit>();
            string path = @"Data\Term.xlsx";
            DataSet dataSetOld = new DataSet("EXCEL");
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=1;MAXSCANROWS=0;HDR=YES;'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSetOld);
            connection.Close();

            for (int i = 0; i < dataSetOld.Tables[0].Rows.Count; i++)
                if (dataSetOld.Tables[0].Rows[i][0].ToString() != "")
                    weights.Add(new Unit("", dataSetOld.Tables[0].Rows[i][1].ToString(), dataSetOld.Tables[0].Rows[i][0].ToString()));
            return weights;
        }

        private static List<string> getData()
        {
            List<string> data = new List<string>();
            for (int i = 0; i < Units.Count; i++)
                data.Add(Units[i].Terminal + "         " + Units[i].Weight);
            data.Sort(new NaturalStringComparer());
            return data;
        }

        private static List<Unit> getAllUnits(Transaction acTrans, PromptSelectionResult selRes)
        {
            List<Unit> units = new List<Unit>();

            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nОшибка выборки блоков!\n");
                acTrans.Commit();
                return null;
            }
            foreach (ObjectId Id in selRes.Value.GetObjectIds())
            {
                BlockTable bt = (BlockTable)acTrans.GetObject(acDb.BlockTableId, OpenMode.ForWrite);
                BlockReference br = (BlockReference)acTrans.GetObject(Id, OpenMode.ForWrite);
                Unit unit = new Unit();
                foreach (ObjectId id in br.AttributeCollection)
                {
                    DBObject obj = id.GetObject(OpenMode.ForRead);
                    AttributeReference attRef = obj as AttributeReference;
                    if (attRef != null)
                    {
                        #region attributes
                        switch (attRef.Tag)
                        {
                            case "НАИМЕНОВАНИЕ":
                                {
                                    unit.Name = attRef.TextString;
                                    break;
                                }
                            case "ОБОЗНАЧЕНИЕ":
                                {
                                    unit.Terminal = attRef.TextString;
                                    break;
                                }
                        }
                        #endregion
                    }
                }
                if (unit.Terminal != null && unit.Name != null)
                {
                    units.Add(unit);
                    br.Color = Color.FromRgb(0, 0, 0);
                }
            }
            return units;
        }
    }
}
