using Microsoft.Win32;
using System.Windows;
using System.IO.Compression;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows.Documents;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace WorkReport
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Png_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog pngFileDialog = new()
            {
                Filter = "ZIP Files (*.zip)|*.zip|All Files (*.*)|*.*",
                Title = "Выберите файлы",
                Multiselect = true, // Разрешаем выбор нескольких файлов
                CheckFileExists = true
            };

            var folderDialog = new CommonOpenFileDialog()
            {
                IsFolderPicker = true,
                Title = "Выберите папку для извлечения"
            };

            // Открываем диалог выбора файлов
            bool? result = pngFileDialog.ShowDialog();

            if (result == true)
            {
                // Получаем массив выбранных файлов
                string[] filePaths = pngFileDialog.FileNames;

                // Открываем диалог выбора папки для извлечения
                if (folderDialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string extractPath = folderDialog.FileName;

                    // Обрабатываем каждый выбранный архив
                    foreach (string filePath in filePaths)
                    {
                        string archiveName = Path.GetFileNameWithoutExtension(filePath) + ".png";
                        Zip_Handler(filePath, extractPath, archiveName);
                    }
                }
                else
                {
                    MessageBox.Show("Действие прервано, давай по новой");
                }
            }
        }

        private void Select_Button_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://web.whatsapp.com/";
            Process.Start(new ProcessStartInfo(url)
            {
                UseShellExecute = true
            });
        }

        private void Zip_Button_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new CommonOpenFileDialog()
            {
                IsFolderPicker = true,
                Title = "Выберите папку с файлами"
            };

            if (folderDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string selectedFolder = folderDialog.FileName;
                var files = Directory.GetFiles(selectedFolder);

                // Находим все Excel-файлы
                var excelFiles = files.Where(f => Path.GetExtension(f).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                                  Path.GetExtension(f).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                                      .ToList();

                // Проходим по каждому Excel файлу
                foreach (var excelFile in excelFiles)
                {
                    string baseName = Path.GetFileNameWithoutExtension(excelFile);

                    // Паттерн для поиска всех фото с таким же базовым именем
                    // Включаем фото без номера и с номерами в скобках
                    string pattern = $"^{Regex.Escape(baseName)}(\\(\\d+\\))?$";

                    // Находим все файлы, которые соответствуют этому паттерну (фото + Excel)
                    var relatedFiles = files.Where(f => Regex.IsMatch(Path.GetFileNameWithoutExtension(f), pattern)).ToList();

                    // ОБЯЗАТЕЛЬНО добавляем Excel файл в список файлов
                    if (!relatedFiles.Contains(excelFile))
                    {
                        relatedFiles.Add(excelFile); // Добавляем сам Excel файл
                    }

                    // Создаем архив, если есть хотя бы один файл
                    if (relatedFiles.Count > 0)
                    {
                        string zipPath = Path.Combine(selectedFolder, baseName + ".zip");
                        CreateArchive(zipPath, relatedFiles.ToArray());
                    }
                }

                MessageBox.Show("Архивация завершена!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Действие прервано, попробуйте снова.");
            }
        }

        private void CreateArchive(string zipPath, string[] filesToAdd)
        {
            using (ZipArchive archive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (string file in filesToAdd)
                {
                    // Убедимся, что мы добавляем файл в архив с правильным именем
                    archive.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
        }



        private void Clear_Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Zip_Handler(string filePath, string extractPath, string archiveName)
        {
            try
            {
                // Удаляем расширение .zip из имени архива
                string archiveNameWithoutExtension = Path.GetFileNameWithoutExtension(archiveName);

                // Создаем временную копию файла
                string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileName(filePath));
                File.Copy(filePath, tempFilePath, overwrite: true);

                using (FileStream zipToOpen = new FileStream(tempFilePath, FileMode.Open))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            // Пропускаем папки
                            if (entry.FullName.EndsWith("/") || entry.FullName.EndsWith("\\"))
                                continue;

                            // Сохраняем структуру папок (если есть)
                            string entryPath = entry.FullName;

                            // Генерируем базовое имя файла (имя архива + расширение файла)
                            string baseFileName = $"{archiveNameWithoutExtension}{Path.GetExtension(entry.Name)}";
                            string destinationPath = Path.Combine(extractPath, baseFileName);

                            // Если файл с таким именем уже существует, добавляем индекс
                            destinationPath = GetUniqueFileName(destinationPath);

                            // Создаем папку, если она не существует
                            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));

                            // Извлекаем файл
                            entry.ExtractToFile(destinationPath, overwrite: true);
                        }
                    }
                }

                // Удаляем временную копию
                File.Delete(tempFilePath);

                MessageBox.Show($"Файлы из архива {archiveName} успешно извлечены!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Файл занят другим процессом: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при извлечении файлов: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Метод для генерации уникального имени файла
        private string GetUniqueFileName(string basePath)
        {
            string directory = Path.GetDirectoryName(basePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(basePath);
            string extension = Path.GetExtension(basePath);

            string uniquePath = basePath;
            int counter = 1;

            // Проверяем, существует ли файл с таким именем
            while (File.Exists(uniquePath))
            {
                // Добавляем индекс к имени файла
                uniquePath = Path.Combine(directory, $"{fileNameWithoutExtension} ({counter}){extension}");
                counter++;
            }

            return uniquePath;
        }

        private void ArchiveCreator(string zipPath, string[] filesToAdd)
        {
           using (ZipArchive newArchive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach(string file in filesToAdd)
                {
                    
                }
            }


           
        }

       
    }
}