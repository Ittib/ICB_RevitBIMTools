using Autodesk.Revit.DB;

namespace MTC_1ClickPlot
{
    public static class Config
    {
        public static readonly string[] DEFAULT_EXCLUSIONS = { "Revision Clouds", "Revision Cloud Tags" };
        public static readonly Color BLACK_COLOR = new Color(0, 0, 0);
        public static readonly string TARGET_PRINTER = "PDF24";
        public static readonly string TEMP_PDF_FOLDER = @"C:\Revit_PDF_Temp";
        public static readonly string TARGET_PARAM_NAME = "FileName";
    }
}
