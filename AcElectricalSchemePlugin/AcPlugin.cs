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
        public void Connection()
        {
            ConnectionScheme.DrawScheme();
        }

        [CommandMethod("Mark", CommandFlags.Session)]
        public void Mark()
        {
            CableMark.CalculateMarks();
        }

        [CommandMethod("Control", CommandFlags.Session)]
        public void Control()
        {
            ControlSchemeClass.drawControlScheme();
        }

        [CommandMethod("Terminals", CommandFlags.Session)]
        public void Terminal()
        {
            TerminalsWeightClass.Calculate();
        }

        [CommandMethod("Article", CommandFlags.Session)]
        public void Article()
        {
            ArticleClass.ChangeArticles();
        }

        [CommandMethod("SLine", CommandFlags.Session)]
        public void SLine()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel|*.xlsx";
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
                SingleLineSchemeClass.drawScheme(openFileDialog1.FileName);
        }
    }
}
