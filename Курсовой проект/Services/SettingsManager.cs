using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Курсовой_проект.Services
{
    public class SettingsManager
    {
        private string _iniPath;

        // Импорт функций WinAPI для работы с INI файлами
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        public SettingsManager()
        {
            // INI файл в папке с программой
            _iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.ini");
        }

        /// <summary>
        /// Сохраняет параметры в INI файл
        /// </summary>
        public void SaveSettings(string function, double x0, double y0, double X, double h)
        {
            WritePrivateProfileString("Parameters", "Function", function, _iniPath);
            WritePrivateProfileString("Parameters", "X0", x0.ToString(System.Globalization.CultureInfo.InvariantCulture), _iniPath);
            WritePrivateProfileString("Parameters", "Y0", y0.ToString(System.Globalization.CultureInfo.InvariantCulture), _iniPath);
            WritePrivateProfileString("Parameters", "X", X.ToString(System.Globalization.CultureInfo.InvariantCulture), _iniPath);
            WritePrivateProfileString("Parameters", "H", h.ToString(System.Globalization.CultureInfo.InvariantCulture), _iniPath);
        }

        /// <summary>
        /// Загружает параметры из INI файла
        /// </summary>
        public (string function, double x0, double y0, double X, double h) LoadSettings()
        {
            string function = ReadValue("Parameters", "Function", "y - x*x + 1");
            double x0 = double.Parse(ReadValue("Parameters", "X0", "0"), System.Globalization.CultureInfo.InvariantCulture);
            double y0 = double.Parse(ReadValue("Parameters", "Y0", "0.5"), System.Globalization.CultureInfo.InvariantCulture);
            double X = double.Parse(ReadValue("Parameters", "X", "2"), System.Globalization.CultureInfo.InvariantCulture);
            double h = double.Parse(ReadValue("Parameters", "H", "0.1"), System.Globalization.CultureInfo.InvariantCulture);

            return (function, x0, y0, X, h);
        }

        private string ReadValue(string section, string key, string defaultValue)
        {
            StringBuilder sb = new StringBuilder(255);
            GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, _iniPath);
            return sb.ToString();
        }
    }
}