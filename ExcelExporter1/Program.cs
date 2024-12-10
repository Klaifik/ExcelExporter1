using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.Util;

public interface IExcelFileProcessor : IDisposable
{
    void DuplicateFile(string sourceFile, string outputFile);
}

public class NpoiExcelFileProcessor : IExcelFileProcessor
{
    private bool _disposed = false;

    public void DuplicateFile(string sourceFile, string outputFile)
    {
        if (!File.Exists(sourceFile))
        {
            throw new FileNotFoundException($"Сурс не найден: {sourceFile}", sourceFile);
        }

        try
        {
            using (var inputStream = File.OpenRead(sourceFile))
            using (var outputStream = File.Create(outputFile))
            {
                // Используем XSSFWorkbook для обработки файлов .xlsx
                IWorkbook workbook = new XSSFWorkbook(inputStream);
                workbook.Write(outputStream);
            }
        }
        catch (IOException ex)
        {
            throw new ExcelProcessingException($"Ошибка ввода-вывода: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            throw new ExcelProcessingException($"Ошибка: {ex.Message}", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Освобождение управляемых ресурсов, если необходимо
            }
            _disposed = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~NpoiExcelFileProcessor()
    {
        Dispose(false);
    }
}

public class ExcelProcessingException : Exception
{
    public ExcelProcessingException(string message, Exception innerException) : base(message, innerException) { }
}

public class Model
{
    public int Number { get; set; }
}

public class ExcelDuplicator : IDisposable
{
    private readonly IExcelFileProcessor _excelProcessor;
    private bool _disposed = false;

    public ExcelDuplicator(IExcelFileProcessor excelProcessor)
    {
        _excelProcessor = excelProcessor ?? throw new ArgumentNullException(nameof(excelProcessor));
    }

    public void Duplicate(IEnumerable<Model> models, string sourceFile, string outputFile)
    {
        ValidateInput(sourceFile, outputFile);

        var errors = new List<string>();
        try
        {
            foreach (var model in models)
            {
                var OutputFile = Path.Combine(Path.GetDirectoryName(outputFile), $"{model.Number}.xlsx");
                _excelProcessor.DuplicateFile(sourceFile, OutputFile);
                Console.WriteLine($"Файл '{OutputFile}' создан.");
            }
        }
        catch (Exception ex)
        {
            errors.Add($"Ошибка создания файла: {ex.Message}");
        }

        if (errors.Any())
        {
            Console.WriteLine("\nОшибка:");
            errors.ForEach(Console.WriteLine);
        }
    }

    private void ValidateInput(string sourceFile, string outputFile)
    {
        if (string.IsNullOrWhiteSpace(sourceFile)) throw new ArgumentNullException(nameof(sourceFile));
        if (string.IsNullOrWhiteSpace(outputFile)) throw new ArgumentNullException(nameof(outputFile));
        if (!File.Exists(sourceFile)) throw new FileNotFoundException("Инпутного файла не существует.", sourceFile);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _excelProcessor?.Dispose();
            }
            _disposed = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~ExcelDuplicator()
    {
        Dispose(false);
    }
}

public class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        string sourceFile = "Data\\input.xlsx"; // Путь к вашему входному файлу

        using (var saveFileDialog = new SaveFileDialog())
        {
            saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить файл как";
            saveFileDialog.FileName = "output.xlsx"; // Имя по умолчанию

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string outputFile = saveFileDialog.FileName;

                var models = new List<Model> { new Model { Number = 1 }, new Model { Number = 2 }, new Model { Number = 3 } };

                try
                {
                    using (var processor = new NpoiExcelFileProcessor())
                    using (var duplicator = new ExcelDuplicator(processor))
                    {
                        duplicator.Duplicate(models, sourceFile, outputFile);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Фатальная ошибка: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Выбор файла отменен.");
            }
        }
    }
}