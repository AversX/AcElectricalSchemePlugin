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
            MainForm mf = new MainForm();
            mf.Show();
        }

        public void Terminate()
        {
        }

        [CommandMethod("AcPluginForm", CommandFlags.Session)]
        public static void AcPluginForm()
        {
            MainForm mf = new MainForm();
            mf.Show();
        }

        [CommandMethod("ConnectionScheme", CommandFlags.Session)]
        public static void ConnectionScheme()
        {
            ConnectionSchemeClass.DrawScheme();
        }

        [CommandMethod("ContourScheme", CommandFlags.Session)]
        public static void ContourScheme()
        {
            ContourSchemeClass.DrawScheme();
        }

        [CommandMethod("ControlSchemeToTxt", CommandFlags.Session)]
        public static void ParseControlSchemeToTxt()
        {
            ContourSchemeClass.ParseControlSchemeToTxt();
        }

        [CommandMethod("Mark", CommandFlags.Session)]
        public static void Mark()
        {
            CableMark.CalculateMarks();
        }

        [CommandMethod("ControlScheme", CommandFlags.Session)]
        public static void ControlScheme()
        {
            ControlSchemeClass.DrawControlScheme();
        }

        [CommandMethod("TerminalsWeight", CommandFlags.Session)]
        public static void TerminalsWeight()
        {
            TerminalsWeightClass.Calculate();
        }

        [CommandMethod("ChangeArticles", CommandFlags.Session)]
        public static void ChangeArticles()
        {
            ArticleClass.ChangeArticles();
        }

        [CommandMethod("SLineScheme", CommandFlags.Session)]
        public static void SLineScheme()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel|*.xlsx";
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
                SingleLineSchemeClass.SLineScheme(openFileDialog1.FileName);
        }

        [CommandMethod("ChangeAttributes", CommandFlags.Session)]
        public static void ChangeAttributes()
        {
            ChangeAttributeClass.FindAndChangeAttributes();
        }

        [CommandMethod("ChangeTextHeght", CommandFlags.Session)]
        public static void ChangeTextHeght()
        {
            TextHeightChangeClass.OpenForm();
        }

        [CommandMethod("MarkExport", CommandFlags.Session)]
        public static void MarkExport()
        {
            ExportTerminalsClass.Export();
        }

        [CommandMethod("Help", CommandFlags.Session)]
        public static void Help()
        {
            HelpForm hForm = new HelpForm();
            hForm.Show();
        }


    }
}
