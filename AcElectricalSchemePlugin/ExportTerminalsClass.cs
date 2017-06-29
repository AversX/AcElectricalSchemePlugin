using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AcElectricalSchemePlugin
{
    static class ExportTerminalsClass
    {
        private static Editor editor;
        private static Document acDoc;
        private static Database acDb;

        struct Catch
        {
            string catchName;
            List<string> terminals;

            public List<string> Terminals { get => terminals; set => terminals = value; }
            public string CatchName { get => catchName; set => catchName = value; }

            public void SortTerminals()
            {
                terminals.Sort(new NaturalStringComparer());
            }
        }

        static public void Export()
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            acDb = acDoc.Database;
            List<string> marks = new List<string>();
            List<string> stickers = new List<string>();
            List<Catch> catches = new List<Catch>();

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    TypedValue[] filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_МАРКИРОВКА,КИА_НАКЛЕЙКИ,КИА_УПОР,КИА_КЛЕММЫ,КИА_ШИРИНА_КЛЕММЫ");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    PromptSelectionResult selRes = acDoc.Editor.GetSelection(filter);
                    LayerTable lt = (LayerTable)acTrans.GetObject(acDb.LayerTableId, OpenMode.ForRead);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\nОшибка выборки!\n");
                        acTrans.Commit();
                    }
                    else
                        foreach (ObjectId Id in selRes.Value.GetObjectIds())
                        {
                            DBText text = (DBText)acTrans.GetObject(Id, OpenMode.ForRead);
                            if (text != null)
                            {
                                LayerTableRecord layer = (LayerTableRecord)acTrans.GetObject(text.LayerId, OpenMode.ForRead);
                                switch (layer.Name)
                                {
                                    case "КИА_МАРКИРОВКА": { marks.Add(text.TextString); break; }
                                    case "КИА_НАКЛЕЙКИ": { stickers.Add(text.TextString); break; }
                                    case "КИА_УПОР":
                                        {
                                            Catch catchUnit = new Catch();
                                            catchUnit.CatchName = text.TextString;
                                            if (catchUnit.CatchName.Contains("K") || catchUnit.CatchName.Contains("К")) catchUnit.Terminals = GetTerminals(acTrans, text.Position, true);
                                            else catchUnit.Terminals = GetTerminals(acTrans, text.Position, false);
                                            catches.Add(catchUnit);
                                            if (!catchUnit.CatchName.Contains("K") && !catchUnit.CatchName.Contains("К")) catches.Add(catchUnit);
                                            break;
                                        }
                                }
                            }
                        }
                }
            }

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                GetData(fbd.SelectedPath + "\\КИА_МАРКИРОВКА", marks);
                GetData(fbd.SelectedPath + "\\КИА_НАКЛЕЙКИ", stickers);
                GetData(fbd.SelectedPath, "\\КИА_УПОР", "\\КИА_КЛЕММЫ", catches);
            }
        }

        
        private static void GetData(string filename, List<string> list)
        {
            list.Sort(new NaturalStringComparer());
            Workbook wBook = null;
            Worksheet sheet = null;
            Excel.Application excel = null;
            try
            {
                excel = new Excel.Application();
                excel.SheetsInNewWorkbook = 1;
                wBook = excel.Workbooks.Add(Type.Missing);
                sheet = (Excel.Worksheet)excel.Sheets.get_Item(1);
                for (int i = 0; i < list.Count; i++)
                {
                    sheet.Cells[1][i + 1] = list[i];
                }
                wBook.SaveAs(filename);
                wBook.Close(true);
                excel.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(excel);
            }
        }

        private static void GetData(string path, string filename1, string filename2, List<Catch> list)
        {
            Workbook wBook = null;
            Worksheet sheet = null;
            Excel.Application excel = null;

            list.Sort(new CatchComparer());
            try
            {
                excel = new Excel.Application();
                excel.SheetsInNewWorkbook = 1;
                wBook = excel.Workbooks.Add(Type.Missing);
                sheet = (Excel.Worksheet)excel.Sheets.get_Item(1);
                for (int i = 0; i < list.Count; i++)
                {
                    sheet.Cells[1][i + 1] = list[i].CatchName;
                }
                wBook.SaveAs(path + filename1);
                wBook.Close(true);
                excel.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(excel);
            }

            try
            {
                excel = new Excel.Application();
                excel.SheetsInNewWorkbook = 1;
                wBook = excel.Workbooks.Add(Type.Missing);
                sheet = (Excel.Worksheet)excel.Sheets.get_Item(1);

                int currentNumOnRow = 0;
                int currentRow = 1;
                int maxNumOnRow = 10;

                for (int i = 0; i < list.Count; i++)
                {
                    list[i].SortTerminals();
                    if (list[i].Terminals.Count <= 10)
                    {
                        if (currentNumOnRow + list[i].Terminals.Count <= maxNumOnRow)
                        {
                            for (int j = 0; j < list[i].Terminals.Count; j++)
                                sheet.Cells[currentNumOnRow + j + 1][currentRow] = list[i].Terminals[j];
                            currentNumOnRow += list[i].Terminals.Count;
                        }
                        else
                        {
                            currentRow++;
                            currentNumOnRow = 0;
                            i--;
                            continue;
                        }
                    }
                    else
                    {
                        if (currentNumOnRow == 0)
                        {
                            int j = 0;
                            while (j < list[i].Terminals.Count)
                            {
                                if (j == 10) { currentRow++; currentNumOnRow = 0; }
                                    sheet.Cells[currentNumOnRow + 1][currentRow] = list[i].Terminals[j];
                                j++;
                                currentNumOnRow++;
                            }
                            currentRow++;
                            currentNumOnRow = 0;
                        }
                        else
                        {
                            currentRow++;
                            currentNumOnRow = 0;
                            int j = 0;
                            while (j < list[i].Terminals.Count)
                            {
                                if (j == 10) { currentRow++; currentNumOnRow = 0; }
                                sheet.Cells[currentNumOnRow + 1][currentRow] = list[i].Terminals[j];
                                j++;
                                currentNumOnRow++;
                            }
                            currentRow++;
                            currentNumOnRow = 0;
                        }
                    }
                }
                wBook.SaveAs(path + filename2);
                wBook.Close(true);
                excel.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(wBook);
                Marshal.ReleaseComObject(excel);
            }
        }

        private static List<string> GetTerminals(Transaction acTrans, Point3d textPos, bool KCatch)
        {
            List<string> list = new List<string>();

            TypedValue[]  filterlist = new TypedValue[2];
            filterlist[0] = new TypedValue((int)DxfCode.Start, "LINE");
            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_ЭЛЕМЕНТЫ");
            SelectionFilter filter = new SelectionFilter(filterlist);

            int numOfLines = 0;
            Point3d[] points = new Point3d[2];
            points[0] = new Point3d(textPos.X, textPos.Y, 0);
            int offPoint = 1;
            while (numOfLines < 2 || offPoint < 3000)
            {
                points[1] = new Point3d(textPos.X + offPoint, textPos.Y, 0);
                Point3dCollection fence = new Point3dCollection(points);
                PromptSelectionResult lineSelRes = acDoc.Editor.SelectFence(fence, filter);
                if (lineSelRes.Status == PromptStatus.OK)
                    if (lineSelRes.Value.Count < 2) numOfLines = 1;
                    else
                    {
                        if (KCatch)
                        {
                            Autodesk.AutoCAD.DatabaseServices.Line lastLine = (Autodesk.AutoCAD.DatabaseServices.Line)acTrans.GetObject(lineSelRes.Value[lineSelRes.Value.Count - 1].ObjectId, OpenMode.ForRead);
                            if (lastLine != null)
                                if (Math.Abs(lastLine.Delta.Y) > 8)
                                {
                                    numOfLines = 2;
                                    points = AreaFromTwoLines(acTrans, lineSelRes);
                                    filterlist = new TypedValue[2];
                                    filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                                    filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_КЛЕММЫ");
                                    filter = new SelectionFilter(filterlist);
                                    PromptSelectionResult terminalsSelRes = acDoc.Editor.SelectCrossingWindow(points[0], points[1], filter);
                                    if (terminalsSelRes.Status == PromptStatus.OK)
                                        foreach (ObjectId Id in terminalsSelRes.Value.GetObjectIds())
                                        {
                                            DBText text = (DBText)acTrans.GetObject(Id, OpenMode.ForRead);
                                            if (text != null) list.Add(text.TextString);
                                        }
                                    break;
                                }
                        }
                        else
                        {
                            numOfLines = 2;
                            points = AreaFromTwoLines(acTrans, lineSelRes);
                            filterlist = new TypedValue[2];
                            filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                            filterlist[1] = new TypedValue((int)DxfCode.LayerName, "КИА_КЛЕММЫ");
                            filter = new SelectionFilter(filterlist);
                            PromptSelectionResult terminalsSelRes = acDoc.Editor.SelectCrossingWindow(points[0], points[1], filter);
                            if (terminalsSelRes.Status == PromptStatus.OK)
                                foreach (ObjectId Id in terminalsSelRes.Value.GetObjectIds())
                                {
                                    DBText text = (DBText)acTrans.GetObject(Id, OpenMode.ForRead);
                                    if (text != null) list.Add(text.TextString);
                                }
                            break;
                        }
                    }
                offPoint++;
            }
            return list;
        }

        private static Point3d[] AreaFromTwoLines(Transaction acTrans, PromptSelectionResult selRes)
        {
            Point3d[] points = new Point3d[2];
            List<Autodesk.AutoCAD.DatabaseServices.Line> lines = new List<Autodesk.AutoCAD.DatabaseServices.Line>();
            foreach (ObjectId Id in selRes.Value.GetObjectIds())
            {
                Autodesk.AutoCAD.DatabaseServices.Line line = (Autodesk.AutoCAD.DatabaseServices.Line)acTrans.GetObject(Id, OpenMode.ForRead);
                if (line != null) lines.Add(line);
            }
            double y = lines[0].StartPoint.Y > lines[0].EndPoint.Y ? lines[0].StartPoint.Y : lines[0].EndPoint.Y;
            points[0] = new Point3d(lines[0].StartPoint.X, y, 0);
            for (int i=1; i<lines.Count; i++)
            {
                double yUp = lines[i].StartPoint.Y > lines[i].EndPoint.Y ? lines[i].StartPoint.Y : lines[i].EndPoint.Y;
                double yDown = lines[i].StartPoint.Y < lines[i].EndPoint.Y ? lines[i].StartPoint.Y : lines[i].EndPoint.Y;
                if (lines[i].StartPoint.X < points[0].X) points[0] = new Point3d(lines[i].StartPoint.X, yUp, 0);
                else points[1] = new Point3d(lines[i].StartPoint.X, yDown, 0);
            }
            return points;
        }


        class CatchComparer : IComparer<Catch>
        {
            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
            static extern int StrCmpLogicalW(string s1, string s2);
            public int Compare(Catch x, Catch y)
            {
                return StrCmpLogicalW(x.CatchName, y.CatchName);
            }
        }
    }
}
