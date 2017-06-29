using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Linq;

namespace AcElectricalSchemePlugin
{
    static class ChangeAttributeClass
    {
        private static Document acDoc;
        private static Editor editor;
        private static List<string> attributes;
        private static SelectionSet selectedBlocks;
        private static BlockTableRecord modSpace;
        private static Transaction acTrans;
        static public void FindAndChangeAttributes()
        {
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database acDb = acDoc.Database;

            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (acTrans = acDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    modSpace = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    attributes = new List<string>();

                    findAttributes();

                    acTrans.Commit();
                }
            }
        }

        static private void findAttributes()
        {
            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "INSERT");
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = editor.GetSelection(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nНе выделено ни одного вхождения блока");
                return;
            }
            else
            {
                selectedBlocks = selRes.Value;
                foreach (ObjectId Id in selectedBlocks.GetObjectIds())
                {
                    BlockReference oEnt = (BlockReference)acTrans.GetObject(Id, OpenMode.ForWrite);
                    foreach (ObjectId id in oEnt.AttributeCollection)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForWrite);
                        AttributeReference attRef = obj as AttributeReference;
                        if ((attRef != null))
                            attributes.Add(attRef.Tag);
                    }
                }
            }
            attributes = attributes.Distinct().ToList();

            ChangeAttributeForm chAttrForm = new ChangeAttributeForm();
            chAttrForm.attrListBox.Items.AddRange(attributes.ToArray());
            chAttrForm.ShowDialog();
        }

        internal static void changeAttributes(string selectedAttribute, string textString)
        {
            foreach (ObjectId Id in selectedBlocks.GetObjectIds())
            {
                BlockReference oEnt = (BlockReference)acTrans.GetObject(Id, OpenMode.ForWrite);
                foreach (ObjectId id in oEnt.AttributeCollection)
                {
                    DBObject obj = id.GetObject(OpenMode.ForWrite);
                    AttributeReference attRef = obj as AttributeReference;
                    if ((attRef != null))
                        if (attRef.Tag == selectedAttribute)
                            attRef.TextString = textString;
                }
            }
        }
    }
}
