using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;

namespace AcElectricalSchemePlugin
{
    static class TextHeightChangeClass
    {
        public static void ChangeTextHeight(bool mTxtBP, bool txtBP, bool mTxtBlcks, bool txtBlcks, bool table, double textHeight, double compCoeff, List<string> selectedBlcks, string fontName)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    acDb.Luprec = 2;

                    TextStyleTable tStyles = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForWrite);
                    ObjectId style;
                    if (tStyles.Has(fontName))
                    {
                        style = tStyles[fontName];
                    }
                    else
                    {
                        style = tStyles["standart"];
                    }
                    if (mTxtBP)
                    {
                        TypedValue[] filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
                        SelectionFilter filter = new SelectionFilter(filterlist);
                        PromptSelectionResult selRes = editor.SelectAll(filter);
                        if (selRes.Status != PromptStatus.OK)
                        {
                            editor.WriteMessage("\nНе найден МТЕКСТ в чертеже");
                        }
                        else
                        {
                            foreach (SelectedObject sObj in selRes.Value)
                            {
                                MText mTxt = (MText)acTrans.GetObject(sObj.ObjectId, OpenMode.ForWrite);
                                if (mTxt.Contents.Contains("лист 5") )
                                    { }
                                mTxt.TextStyleId = style;
                                mTxt.TextHeight = textHeight;
                                //mTxt.Contents = "{\\W" + compCoeff.ToString().Replace(',', '.') + ";" + mTxt.Text +"}";
                                //mTxt.Contents = "\\W" + compCoeff.ToString().Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator) + ";" + mTxt.Text;
                                mTxt.Contents = "\\pxi-" + compCoeff.ToString().Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator) + ",l1;" + mTxt.Text;
                            }
                        }
                    }
                    if (txtBP)
                    {
                        TypedValue[] filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                        SelectionFilter filter = new SelectionFilter(filterlist);
                        PromptSelectionResult selRes = editor.SelectAll(filter);
                        if (selRes.Status != PromptStatus.OK)
                        {
                            editor.WriteMessage("\nНе найден ТЕКСТ в чертеже");
                        }
                        else
                        {
                            foreach (SelectedObject sObj in selRes.Value)
                            {
                                DBText txt = (DBText)acTrans.GetObject(sObj.ObjectId, OpenMode.ForWrite);
                                txt.TextStyleId = style;
                                txt.Height = textHeight;
                                txt.WidthFactor = compCoeff; 
                            }
                        }
                    }
                    if (table)
                    {
                        TypedValue[] filterlist = new TypedValue[1];
                        filterlist[0] = new TypedValue((int)DxfCode.Start, "ACAD_TABLE");
                        SelectionFilter filter = new SelectionFilter(filterlist);
                        PromptSelectionResult selRes = editor.SelectAll(filter);
                        if (selRes.Status != PromptStatus.OK)
                        {
                            editor.WriteMessage("\nНе найдены таблицы в чертеже");
                        }
                        else
                        {
                            foreach (SelectedObject sObj in selRes.Value)
                            {
                                Table tbl = (Table)acTrans.GetObject(sObj.ObjectId, OpenMode.ForWrite);
                                tbl.Cells.Style = "Данные";
                                tbl.Cells.TextStyleId = style;
                                tbl.Cells.TextHeight = textHeight;
                            }
                        }
                    }
                    if (selectedBlcks.Count>0)
                    {
                        BlockTable bt = (BlockTable)acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead);
                        for (int i = 0; i < selectedBlcks.Count; i++)
                        {
                            if (bt.Has(selectedBlcks[i]))
                            {
                                BlockTableRecord btr = acTrans.GetObject(bt[selectedBlcks[i]], OpenMode.ForRead) as BlockTableRecord;
                                foreach (ObjectId Id in btr.Cast<ObjectId>())
                                {
                                    if (mTxtBlcks)
                                    {
                                        MText mTxt = acTrans.GetObject(Id, OpenMode.ForWrite) as MText;
                                        if (mTxt != null)
                                        {
                                            mTxt.TextStyleId = style;
                                            mTxt.TextHeight = textHeight;
                                            mTxt.Contents = "{\\W" + compCoeff.ToString().Replace(',', '.') + ";" + mTxt.Text;
                                        }
                                    }
                                    if (txtBlcks)
                                    {
                                        DBText txt = acTrans.GetObject(Id, OpenMode.ForWrite) as DBText;
                                        if (txt != null)
                                        {
                                            txt.TextStyleId = style;
                                            txt.Height = textHeight;
                                            txt.WidthFactor = compCoeff;
                                        }
                                    }
                                }
                            }
                        }
                    }
                   // acDb.Audit(true, true);
                    acTrans.Commit();
                }
            }       
        }

        public static void OpenForm()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;
            List<string> namesOfBlocks = new List<string>();
            List<string> textStyles = new List<string>();
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId Id in bt)
                    {
                        BlockTableRecord btr = acTrans.GetObject(Id, OpenMode.ForRead) as BlockTableRecord;
                        if (!btr.Name.ToUpper().Contains("MODEL_SPACE") && !btr.Name.ToUpper().Contains("PAPER_SPACE")) namesOfBlocks.Add(btr.Name);
                    }
                    TextStyleTable tst = (TextStyleTable)acTrans.GetObject(acDb.TextStyleTableId, OpenMode.ForRead);
                    foreach (ObjectId Id in tst)
                    {
                        TextStyleTableRecord tstr = acTrans.GetObject(Id, OpenMode.ForRead) as TextStyleTableRecord;
                        textStyles.Add(tstr.Name);
                    }
                    acTrans.Commit();
                }
            }
            ChangeTextHeightForm chTxtHForm = new ChangeTextHeightForm(namesOfBlocks, textStyles);
            chTxtHForm.Show();
        }
    }
}
