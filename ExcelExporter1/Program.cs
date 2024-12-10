using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public interface IExcelFileProcessor : IDisposable
{
    void DuplicateFile(string sourceFile, string destinationFile);
}

public class NpoiExcelFileProcessor : IExcelFileProcessor
{
    private bool _disposed = false;
    private Stream inputStream;

    public void DuplicateFile(string sourceFile, string destinationFile)
    {
        if (!File.Exists(sourceFile))
        {
            throw new FileNotFoundException($"Сурс не найден: {sourceFile}", sourceFile);
        }

        try
        {
            var workbook = WorkbookFactory.Create(inputStream);
            using (var inputStream = File.OpenRead(sourceFile))
            using (var outputStream = File.Create(destinationFile))
            {
                workbook.Write(outputStream);
            }
        }
        catch (IOException ex)
        {
            throw new ExcelProcessingException($"Ошибка: {ex.Message}", ex);
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
                if (inputStream != null)
                {
                    inputStream.Dispose();
                }
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
    private bool _disposed = true;

    public ExcelDuplicator(IExcelFileProcessor excelProcessor)
    {
        _excelProcessor = excelProcessor;
    }

    public void Duplicate(IEnumerable<Model> models, string inputFile, string outputDirectory)
    {
        ValidateInput(inputFile, outputDirectory);

        Directory.CreateDirectory(outputDirectory);

        var errors = new List<string>();
        foreach (var model in models)
        {
            var outputFile = Path.Combine(outputDirectory, $"{model.Number}.xlsx");
            try
            {
                _excelProcessor.DuplicateFile(inputFile, outputFile);
                Console.WriteLine($"Файл '{outputFile}' Создан.");
            }
            catch (Exception ex)
            {
                errors.Add($"Ошибка создания файла '{outputFile}': {ex.Message}");
            }
        }

        if (errors.Any())
        {
            Console.WriteLine("\nОшибка:");
            errors.ForEach(Console.WriteLine);
        }
    }

    private void ValidateInput(string inputFile, string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(inputFile)) throw new ArgumentNullException(nameof(inputFile));
        if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentNullException(nameof(outputDirectory));
        if (!File.Exists(inputFile)) throw new FileNotFoundException("Инпутного файла не существует.", inputFile);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing && _excelProcessor != null)
            {
                _excelProcessor.Dispose();
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
        using (var folderBrowserDialog = new FolderBrowserDialog())
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string outputDirectory = folderBrowserDialog.SelectedPath;

                var models = new List<Model> { new Model { Number = 1 }, new Model { Number = 2 }, new Model { Number = 3 } };

                string inputFile = "input.xlsx";

                try
                {
                    using (var processor = new NpoiExcelFileProcessor())
                    using (var duplicator = new ExcelDuplicator(processor))
                    {
                        duplicator.Duplicate(models, inputFile, outputDirectory);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Фаталити: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Выбор папки отменен.");
            }
        }
    }
}