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

//using System.Data.OleDb;
//using System.Data;

namespace AcElectricalSchemePlugin
{
    public class AcPlugin : IExtensionApplication
    {
        public Autodesk.AutoCAD.ApplicationServices.Application acadApp;
        public Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
        public void Initialize()
        {
            editor.WriteMessage("Плагин успешно загружен.");
        }

        public void Terminate()
        {
        }

        [CommandMethod("Scheme", CommandFlags.Session)]
        public void Start()
        {
            ConnectionScheme.DrawScheme();
        }

        [CommandMethod("Mark", CommandFlags.Session)]
        public void Mark()
        {
            CableMark.CalculateMarks();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text|*.txt";
            saveFileDialog1.ShowDialog();
            if(saveFileDialog1.FileName != "")
                System.IO.File.WriteAllLines(saveFileDialog1.FileName, CableMark.getData());
        }

        [CommandMethod("Table", CommandFlags.Session)]
        public void Table()
        {
            List<Table> tt = new List<Table>();
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acModSpace;
                    acModSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    List<Table> t = new List<Table>();

                    for (int i = 0; i < 3; i++)
                    {
                        Table table = new Table();
                        table.TableStyle = acDb.Tablestyle;
                        table.Position = new Point3d(80*i, 0, 0);
                        table.SetSize(4, 1);

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

                        table.Columns[0].Width = 30;
                        table.GenerateLayout();
                        t.Add(table);

                        acModSpace.AppendEntity(table);
                        acTrans.AddNewlyCreatedDBObject(table, true);
                    }

                    for (int i = 0; i < 3; i++)
                    {
                        t[i].InsertColumns(t[i].Columns.Count, 20, 1);
                    }
                    acTrans.Commit();
                }
            }
        }
    }
}
