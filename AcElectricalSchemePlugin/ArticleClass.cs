using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.EditorInput;

namespace AcElectricalSchemePlugin
{
    static class ArticleClass
    {
        private static Document acDoc;
        private static Editor editor;
        private static List<preUnit> oldUnits;
        private static List<preUnit> newUnits;
        private static List<Unit> units;
        private static string pathNew = "";
        private static string pathOld = ""; 

        struct preUnit
        {
            public string Number;
            public string Article;
            public preUnit(string number, string article)
            {
                Number = number;
                Article = article;
            }
        }

        struct Unit
        {
            public string Number;
            public string Article;
            public string OldNumber;
            public Unit(string number, string article, string oldNumber)
            {
                Number = number;
                OldNumber = oldNumber;
                Article = article;
            }
        }

        public static void ChangeArticles()
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            oldUnits = new List<preUnit>();
            newUnits = new List<preUnit>();
            units = new List<Unit>();
            pathNew = "";
            pathOld = ""; 

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "(*.xlsx)|*.xlsx";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                pathNew = file.FileName;

            file = new OpenFileDialog();
            file.Filter = "(*.xlsx)|*.xlsx";
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                pathOld = file.FileName;

            if (pathNew != "" && pathOld != "")
            {
                LoadAndParseData();
            }

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    List<DBText> textBlocks = new List<DBText>();
                    TypedValue[] filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "SP_NUM");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    PromptSelectionResult selRes = editor.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nОшибка выборки текстовых полей!\n");
                        acTrans.Commit();
                        return;
                    }
                    foreach (ObjectId id in selRes.Value.GetObjectIds())
                    {
                        textBlocks.Add((DBText)acTrans.GetObject(id, OpenMode.ForWrite));
                    }
                    for (int i = 0; i < units.Count; i++)
                    {
                        List<DBText> txt = textBlocks.FindAll(x => x.TextString.ToUpper() == units[i].OldNumber.ToUpper());
                        if (txt.Count>0)
                        {
                            for (int j = 0; j < txt.Count; j++)
                            {
                                txt[j].TextString = units[i].Number;
                                txt[j].Color = Color.FromRgb(0, 0, 0);
                                textBlocks.Remove(txt[j]);
                            }
                            units.RemoveAt(i);
                            i--;
                        }
                        else
                        {
                            editor.WriteMessage("В чертеже не найдено текстовое поле с номером \"{0}\"", units[i].OldNumber);
                        }
                    }
                    SaveFileDialog f = new SaveFileDialog();
                    if (f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        List<string> text = new List<string>();
                        text.Add("Артикулы старой спецификации не найденные в новой");
                        for (int i = 0; i < oldUnits.Count; i++)
                            text.Add(string.Format("No.: {0}     Art.:{1}", oldUnits[i].Number, oldUnits[i].Article));

                        text.Add("Незадействованные артикулы новой спецификации");
                        for (int i = 0; i < newUnits.Count; i++)
                            text.Add(string.Format("No.: {0}     Art.:{1}", newUnits[i].Number, newUnits[i].Article));

                        text.Add("Номера не найденные в чертеже");
                        for (int i = 0; i < units.Count; i++)
                            text.Add(string.Format("oldNo.: {0}     newNo.: {1}     Art.:{2}", units[i].OldNumber, units[i].Number, units[i].Article));

                        text.Add("Незадействованные номера в документе");
                        for (int i = 0; i < textBlocks.Count; i++)
                            text.Add(string.Format("No.: {0}", textBlocks[i].TextString));

                        File.WriteAllLines(f.FileName+".txt", text);
                    }
                    acTrans.Commit();
                }
            }
        }

        private static void LoadAndParseData()
        {
            DataSet dataSetOld = new DataSet("EXCEL");
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathOld + ";Extended Properties='Excel 12.0;IMEX=1;MAXSCANROWS=0;HDR=NO;'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            //adapter.FillSchema(dataSetOld, SchemaType.Source);
            adapter.Fill(dataSetOld);
            connection.Close();

            for(int i=0; i<dataSetOld.Tables[0].Rows.Count; i++)
            {
                oldUnits.Add(new preUnit(dataSetOld.Tables[0].Rows[i][0].ToString(), dataSetOld.Tables[0].Rows[i][2].ToString()));
            }

            DataSet dataSetNew = new DataSet("EXCEL");
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathNew + ";Extended Properties='Excel 12.0;IMEX=1;MAXSCANROWS=0;HDR=NO;'";
            connection = new OleDbConnection(connectionString);
            connection.Open();

            schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            select = String.Format("SELECT * FROM [{0}]", sheet1);
            adapter = new OleDbDataAdapter(select, connection);
            //adapter.FillSchema(dataSetNew, SchemaType.Source);
            adapter.Fill(dataSetNew);
            connection.Close();

            for (int i = 0; i < dataSetNew.Tables[0].Rows.Count; i++)
            {
                newUnits.Add(new preUnit(dataSetNew.Tables[0].Rows[i][0].ToString(), dataSetNew.Tables[0].Rows[i][2].ToString()));
            }

            for(int i=0; i<oldUnits.Count; i++)
            {
                int index = newUnits.FindIndex(x => x.Article.ToUpper() == oldUnits[i].Article.ToUpper());
                if (index != -1)
                {
                    units.Add(new Unit(newUnits[index].Number, newUnits[index].Article, oldUnits[i].Number));
                    newUnits.RemoveAt(index);
                    oldUnits.RemoveAt(i);
                    i--;
                }
            }
        }
    }
}
