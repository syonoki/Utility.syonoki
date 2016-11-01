using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Utility.syonoki.MSOffice {
    public class ExcelSimpleApi {
        private readonly string currentDirectory_ = Directory.GetCurrentDirectory() + "\\";
        public Application application { get; } = new Application();

        public void quit() {
            application.Quit();
        }

        public Workbook fileopen(string filePath, bool visibility = true) {
            application.Visible = visibility;
            Workbook wkb;

            try {
                wkb = openWorkbook(filePath);
            }
            catch (FileNotFoundException) {
                throw new FileNotFoundException("파일을 찾을 수 없습니다.");
            }

            return wkb;
        }

        private Workbook openWorkbook(string filePath) {
            bool hasFullPath = filePath.Contains("\\");
            var path = hasFullPath
                ? filePath
                : currentDirectory_ + "\\" + filePath;

            return application.Workbooks.Open(path);
        }

        public static Application activeApplication() {
            return (Application) Marshal.GetActiveObject("Excel.Application");
        }
    }
}