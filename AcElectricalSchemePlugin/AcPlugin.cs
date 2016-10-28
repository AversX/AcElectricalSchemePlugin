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

        [CommandMethod("Start", CommandFlags.Session)]
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
    }
}
