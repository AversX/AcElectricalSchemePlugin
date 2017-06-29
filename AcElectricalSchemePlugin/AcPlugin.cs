using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows.Forms;

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

        [CommandMethod("Contour", CommandFlags.Session)]
        public void Contour()
        {
            ContourSchemeClass.DrawScheme();
        }

        [CommandMethod("ControlToTxt", CommandFlags.Session)]
        public void ParseControlToTxt()
        {
            ContourSchemeClass.ParseControlSchemeToTxt();
        }

        [CommandMethod("Mark", CommandFlags.Session)]
        public void Mark()
        {
            CableMark.CalculateMarks();
        }

        [CommandMethod("Control", CommandFlags.Session)]
        public void Control()
        {
            ControlSchemeClass.DrawControlScheme();
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

        [CommandMethod("Help", CommandFlags.Session)]
        public void Help()
        {
            HelpForm hForm = new HelpForm();
            hForm.Show();
        }

        [CommandMethod("ChangeAttribute", CommandFlags.Session)]
        public void ChangeAttribute()
        {
            ChangeAttributeClass.FindAndChangeAttributes();
        }

        [CommandMethod("MarkExport", CommandFlags.Session)]
        public void MarkExport()
        {
            ExportTerminalsClass.Export();
        }
    }
}
