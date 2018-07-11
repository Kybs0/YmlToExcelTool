using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Aspose.Cells;
using Microsoft.Win32;
using Path = System.IO.Path;
using SaveOptions = System.Xml.Linq.SaveOptions;
using Style = Aspose.Cells.Style;

namespace YmlToExcelTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<SourceFileModel> SourceYmlList = new List<SourceFileModel>() { };
        public List<string> SourceExcelList = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void AddYmlFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            string tag = (sender as Button)?.Tag.ToString();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.Yml文件|*.yml";
            if (openFileDialog.ShowDialog(this) == true)
            {
                string filePath = openFileDialog.FileName;
                if (File.Exists(filePath))
                {
                    SourceYmlList.Add(new SourceFileModel()
                    {
                        LanguageType = tag,
                        FilePath = filePath
                    });


                    switch (tag)
                    {
                        case LanguageType.EnUs:
                            {
                                FileEnglishListItemsControl.ItemsSource =
                                    SourceYmlList.Where(i => i.LanguageType == LanguageType.EnUs);
                            }
                            break;
                        case LanguageType.ZhCHS:
                            {
                                ChineseFileListItemsControl.ItemsSource =
                                    SourceYmlList.Where(i => i.LanguageType == LanguageType.ZhCHS);
                            }
                            break;
                    }
                }
            }
        }
        private void AddExcelFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.Excel文件|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog(this) == true)
            {
                string filePath = openFileDialog.FileName;
                if (File.Exists(filePath))
                {
                    SourceExcelList.Add(filePath);
                    ExcelItemsControl.ItemsSource = SourceExcelList;
                }
            }
        }

        private void TranslateButton_OnClick(object sender, RoutedEventArgs e)
        {
            List<LangModel> chineseData = ReadYmlData(SourceYmlList.Where(i => i.LanguageType == LanguageType.ZhCHS).Select(i => i.FilePath).ToList());
            List<LangModel> englishData = ReadYmlData(SourceYmlList.Where(i => i.LanguageType == LanguageType.EnUs).Select(i => i.FilePath).ToList());

            var excelData = ReadExcelData(SourceExcelList);
            chineseData.AddRange(excelData.Where(i => i.LanguageType == LanguageType.ZhCHS));
            englishData.AddRange(excelData.Where(i => i.LanguageType == LanguageType.EnUs));
            if (TranslationComboBox.Text == "转出到Excel")
            {
                ExportDataToExcel(chineseData, englishData);
            }
            else
            {
                SaveFileDialog sfDialog = new SaveFileDialog();
                sfDialog.InitialDirectory = @"C:\Users\user\Desktop\";
                sfDialog.Filter = "*.Yml文件|*.yml";
                if ((sfDialog.ShowDialog() == true))
                {
                    ExportDataToYml(chineseData, sfDialog.FileName + "_zh-CHS");
                    ExportDataToYml(englishData, sfDialog.FileName + "_en-US");
                }
            }
        }

        #region 导出到YML文件
        /// <summary>
        /// 读取多个Excel文件
        /// </summary>
        /// <param name="fileList"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        private List<LangModel> ReadExcelData(List<string> fileList, int columnIndex = 1)
        {
            List<LangModel> langData = new List<LangModel>();
            foreach (var excelFile in fileList)
            {
                var workbook = new Workbook(excelFile);
                var sheet = workbook.Worksheets[0];
                for (int i = 1; i < sheet.Cells.MaxDataRow; i++)
                {
                    string key = sheet.Cells[i, 0].Value.ToString();
                    if (sheet.Cells[i, columnIndex].Value == null || string.IsNullOrEmpty(sheet.Cells[i, columnIndex].Value.ToString().Trim()))
                    {
                        continue;
                    }
                    string chineseValue = sheet.Cells[i, columnIndex]?.Value?.ToString();
                    langData.Add(new LangModel() { LanguageType = LanguageType.ZhCHS, Key = key, Value = chineseValue?.Trim(), FileName = Path.GetFileNameWithoutExtension(excelFile) });
                    string englishValue = sheet.Cells[i, columnIndex + 1]?.Value?.ToString();
                    langData.Add(new LangModel() { LanguageType = LanguageType.EnUs, Key = key, Value = englishValue?.Trim(), FileName = Path.GetFileNameWithoutExtension(excelFile) });
                }
            }

            return langData;
        }

        /// <summary>
        /// 导出数据到Yml文件
        /// </summary>
        /// <param name="langData"></param>
        /// <param name="sfDialogFileName"></param>
        private void ExportDataToYml(List<LangModel> langData, string sfDialogFileName)
        {
            langData = langData.OrderBy(i => i.Key).ToList();

            List<string> contentList = new List<string>();
            List<KeyModel> keyList = new List<KeyModel>();
            foreach (var langModel in langData)
            {
                string[] keys = langModel.Key.Replace("Lang.", "").Split('.');
                for (int i = 0; i < keys.Length; i++)
                {
                    if (i + 1 <= keyList.Count && keyList[i].Key == keys[i])
                    {
                        continue;
                    }
                    keyList.RemoveRange(i, keyList.Count - i);
                    keyList.Add(new KeyModel() { Key = keys[i] });
                    string value = langData.First(p => p.Key == langModel.Key).Value ?? "";

                    if (value.Contains("\\n") || value.Contains("\r\n") || value.Contains("\\"))
                    {
                        value = "\"" + value.Replace("\r\n", "\\n") + "\"";
                    }
                    if (string.IsNullOrEmpty(value))
                    {
                        value = "NoneValue";
                    }

                    var content = keys.Length == i + 1 ? GenerateContent(i * 2, keys[i] + ": " + value) : GenerateContent(i * 2, keys[i] + ":");

                    contentList.Add(content);
                }
            }

            File.WriteAllLines(sfDialogFileName, contentList, Encoding.UTF8);
        }

        private string GenerateContent(int i, string keyValue)
        {
            string content = string.Empty;
            for (int j = 0; j < i; j++)
            {
                content += " ";
            }
            content += keyValue;
            return content;
        }
        #endregion

        #region 导出到Excel文件
        /// <summary>
        /// 获取多个Yml文件的数据
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        private List<LangModel> ReadYmlData(List<string> fileList)
        {
            var lineList = new List<YmlLineModel>();
            foreach (var ymlFile in fileList)
            {
                lineList.AddRange(File.ReadAllLines(ymlFile, Encoding.UTF8).Select(i => new YmlLineModel() { FileName = Path.GetFileNameWithoutExtension(ymlFile), LineText = i }).ToList());
            }

            var valueList = new List<LangModel>();

            var keyList = new List<KeyModel>() { new KeyModel() { SpaceCount = -1, Key = "Lang" } };
            foreach (var line in lineList)
            {
                if (!line.LineText.Contains(":")) continue;

                string lineKey = line.LineText.Substring(0, line.LineText.IndexOf(":", StringComparison.Ordinal));
                string linevalue = line.LineText.Substring(line.LineText.IndexOf(":", StringComparison.Ordinal) + 1);

                int spaceCount = (lineKey.Length - lineKey.Replace(" ", "").Length);

                int keyCount = keyList.Count;
                var currentKey = String.Empty;
                for (int i = 0; i < keyCount + 1; i++)
                {
                    if (i < keyCount && keyList[i].SpaceCount < spaceCount)
                    {
                        currentKey += keyList[i].Key + ".";
                    }
                    else
                    {
                        if (i < keyCount)
                        {
                            keyList.RemoveRange(i, keyCount - i);
                        }
                        keyList.Add(new KeyModel() { SpaceCount = spaceCount, Key = lineKey.Trim() });
                        if (!string.IsNullOrEmpty(linevalue))
                        {
                            currentKey += lineKey.Trim();
                            valueList.Add(new LangModel() { Key = currentKey, Value = linevalue.Trim(), FileName = line.FileName });
                        }
                        break;
                    }
                }
            }

            return valueList;
        }
        /// <summary>
        /// 导出数据到Excel
        /// </summary>
        /// <param name="soureChineseData"></param>
        /// <param name="sourceEnglishData"></param>
        private void ExportDataToExcel(List<LangModel> soureChineseData, List<LangModel> sourceEnglishData)
        {
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Style style = new Style();
            style.Number = 49;
            style.IsTextWrapped = true;

            sheet.Cells[0, 0].Value = "Key";
            sheet.Cells[0, 1].Value = LanguageType.ZhCHS;
            sheet.Cells.ApplyColumnStyle(1, style, new StyleFlag());
            sheet.Cells[0, 2].Value = LanguageType.EnUs;
            sheet.Cells.ApplyColumnStyle(2, style, new StyleFlag());
            sheet.Cells[0, 3].Value = "中文长度<英文长度";
            sheet.Cells.ApplyColumnStyle(3, style, new StyleFlag());

            int rowIndex = 1;
            soureChineseData = soureChineseData.Where(i => !string.IsNullOrWhiteSpace(i.Value)).ToList();
            foreach (var langSource in soureChineseData)
            {
                sheet.Cells[rowIndex, 0].Value = langSource.Key;
                sheet.Cells[rowIndex, 1].Value = langSource.Value;
                if (sourceEnglishData.Any(i => i.Key == langSource.Key))
                {
                    var enValue = sourceEnglishData.First(i => i.Key == langSource.Key).Value;
                    sheet.Cells[rowIndex, 2].Value = enValue;

                    //中英文长度对比
                    if (TextHelper.CompareTextLength(langSource.Value, enValue))
                    {
                        sheet.Cells[rowIndex, 3].Value = "True";
                    }
                }
                rowIndex++;
            }

            //调整列表显示
            sheet.FreezePanes(1, 1, 1, 0);
            sheet.AutoFitColumns();
            sheet.Cells.SetColumnWidth(1, 40);
            sheet.Cells.SetColumnWidth(1, 70);
            sheet.Cells.SetColumnWidth(2, 70);

            SaveFileDialog sfDialog = new SaveFileDialog();
            sfDialog.InitialDirectory = @"C:\Users\user\Desktop\";
            sfDialog.Filter = "*.Excel文件|*.xlsx";
            if (sfDialog.ShowDialog() == true)
            {
                string filePath = sfDialog.FileName;
                wb.Save(filePath);
            }
        }
        #endregion
    }

    public class YmlLineModel
    {
        public string LineIndex { get; set; }
        public string LineText { get; set; }
        public string FileName { get; set; }
    }
    public class KeyModel
    {
        public int SpaceCount { get; set; }
        public string Key { get; set; }
    }
    public class LangModel
    {
        public string LanguageType { get; set; }
        /// <summary>
        /// 字段所属文件名
        /// </summary>
        public string FileName { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class SourceFileModel
    {
        public string LanguageType { get; set; }

        public string FilePath { get; set; }
    }

    public static class LanguageType
    {
        public const string ZhCHS = "zhCHS";
        public const string EnUs = "enUs";
    }
}
