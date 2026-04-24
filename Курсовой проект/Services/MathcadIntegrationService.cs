using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.IO;
using System.Diagnostics;

namespace Курсовой_проект.Services
{
    public class MathcadIntegrationService
    {
        private dynamic mathcadApp;
        private dynamic worksheet;

        private string GetTemplatePath()
        {
            string[] possiblePaths = new string[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "MathcadTemplate.mctx"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MathcadTemplate.mctx"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Templates", "MathcadTemplate.mctx"),
            };

            foreach (string path in possiblePaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    Debug.WriteLine($"Шаблон найден: {fullPath}");
                    return fullPath;
                }
            }
            return null;
        }

        public bool StartMathcad()
        {
            try
            {
                string templatePath = GetTemplatePath();
                if (templatePath == null)
                {
                    MessageBox.Show("Шаблон Mathcad не найден!", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                Type mathcadType = Type.GetTypeFromProgID("MathcadPrime.Application.10");
                if (mathcadType == null)
                {
                    mathcadType = Type.GetTypeFromProgID("MathcadPrime.Application");
                }
                if (mathcadType == null)
                {
                    MessageBox.Show("Mathcad Prime 10 не найден в системе", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                mathcadApp = Activator.CreateInstance(mathcadType);
                mathcadApp.Visible = true;
                worksheet = mathcadApp.Open(templatePath);

                if (worksheet == null)
                {
                    MessageBox.Show("Не удалось открыть шаблон Mathcad", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка запуска Mathcad: {ex.Message}", "Ошибка");
                return false;
            }
        }

        public void SetRealValue(string aliasName, double value)
        {
            try
            {
                worksheet.SetRealValue(aliasName, value, "");
                Debug.WriteLine($"✓ Передано {aliasName} = {value}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetRealValue error for {aliasName}: {ex.Message}");
            }
        }

        public void SetStringValue(string aliasName, string value)
        {
            try
            {
                worksheet.SetStringValue(aliasName, value);
                Debug.WriteLine($"✓ Передана строка {aliasName} = {value}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetStringValue error for {aliasName}: {ex.Message}");
                MessageBox.Show($"Ошибка передачи строки: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public void CloseMathcad()
        {
            try
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (mathcadApp != null) mathcadApp.Quit();
            }
            catch { }
        }
    }
}