using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

// Change this to match your program's normal namespace
namespace Utility.syonoki
{
    public class IniFile   // revision 10
    {
        string path_;
        string exe_ = Assembly.GetExecutingAssembly().GetName().Name;

        [DllImport("kernel32")]
        static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(string section, string key, string Default, StringBuilder retVal, int size, string filePath);

        public IniFile(string iniPath = null)
        {
            path_ = new FileInfo(iniPath ?? exe_ + ".ini").FullName;
        }

        public string read(string key, string section = null)
        {
            var retVal = new StringBuilder(255);
            GetPrivateProfileString(section ?? exe_, key, "", retVal, 255, path_);
            return retVal.ToString();
        }

        public void write(string key, string value, string section = null)
        {
            WritePrivateProfileString(section ?? exe_, key, value, path_);
        }

        public void deleteKey(string key, string section = null)
        {
            write(key, null, section ?? exe_);
        }

        public void deleteSection(string section = null)
        {
            write(null, null, section ?? exe_);
        }

        public bool keyExists(string key, string section = null)
        {
            return read(key, section).Length > 0;
        }
    }
}