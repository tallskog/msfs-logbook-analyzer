using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Text.Json;
using System.Xml.Serialization;
using Microsoft.Win32;

using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using System.Runtime.InteropServices.WindowsRuntime;

namespace MsfsLogbookAnalyzer
{
    public partial class MainWindow : Window
    {
        private List<FlightRecord> _flightRecords = new List<FlightRecord>();

        public MainWindow()
        {
            InitializeComponent();
        }

        // �️ Process screenshot button (paste + OCR)
        private async void ProcessScreenshot_Click(object sender, RoutedEventArgs e)
        {
            // Paste
            if (Clipboard.ContainsImage())
            {
                BitmapSource image = Clipboard.GetImage();
                MyImageControl.Source = image;
            }
            else
            {
                MessageBox.Show("No image in clipboard.");
                return;
            }

            OutputTextBox.Text = "Running OCR...\n";

            var result = await RunOCR((BitmapSource)MyImageControl.Source);

            var records = BuildRows(result);

            OutputTextBox.Text = string.Empty;

            if (records.Count == 0)
            {
                OutputTextBox.Text = "No flight records found.\n\nOCR raw lines:\n" + GetRawOcrText(result) + "\n\nRow diagnostics:\n" + GetOcrDebugText(result);
                return;
            }

            _flightRecords.AddRange(records);

            foreach (var r in records)
            {
                OutputTextBox.AppendText(
                    $"{r.Date} | {r.Aircraft} | {r.From} → {r.To} | {r.Duration}\n");
            }

            UpdateRouteCount();
        }

        private void UpdateRouteCount()
        {
            RouteCountTextBlock.Text = $"Routes collected: {_flightRecords.Distinct().Count()}";
        }

        // 🧠 OCR Engine
        private async Task<OcrResult> RunOCR(BitmapSource bitmap)
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));

            using var stream = new MemoryStream();
            encoder.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);

            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream.AsRandomAccessStream());
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync();

            var ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
            return await ocrEngine.RecognizeAsync(softwareBitmap);
        }

        // 🧱 Build rows from OCR (KEY PART)
        private List<FlightRecord> BuildRows(OcrResult result)
        {
            var words = result.Lines
                .SelectMany(l => l.Words)
                .Select(w => new WordBox
                {
                    Text = NormalizeText(w.Text),
                    X = w.BoundingRect.X,
                    Y = w.BoundingRect.Y
                })
                .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                .OrderBy(w => w.Y)
                .ThenBy(w => w.X)
                .ToList();

            var rows = new List<List<WordBox>>();
            const double threshold = 12.0;

            foreach (var word in words)
            {
                var row = rows
                    .OrderBy(r => Math.Abs(r.Average(w => w.Y) - word.Y))
                    .FirstOrDefault(r => Math.Abs(r.Average(w => w.Y) - word.Y) < threshold);

                if (row == null)
                {
                    rows.Add(new List<WordBox> { word });
                }
                else
                {
                    row.Add(word);
                }
            }

            foreach (var row in rows)
            {
                row.Sort((a, b) => a.X.CompareTo(b.X));
            }

            var records = new List<FlightRecord>();

            foreach (var row in rows)
            {
                var tokens = row
                    .Select(w => w.Text ?? string.Empty)
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .ToList();

                if (tokens.Count == 0)
                    continue;

                var record = ParseRow(tokens);

                if (record != null)
                    records.Add(record);
            }

            return records;
        }

        private string GetRawOcrText(OcrResult result)
        {
            return string.Join("\n", result.Lines.Select(line => string.Join(" ", line.Words.Select(w => w.Text))));
        }

        private string GetOcrDebugText(OcrResult result)
        {
            var rows = GetRowTokenGroups(result);
            var lines = new List<string>();

            for (int index = 0; index < rows.Count; index++)
            {
                var tokens = rows[index];
                var record = ParseRow(tokens);
                var rowText = string.Join(" | ", tokens);

                if (record != null)
                {
                    lines.Add($"Row {index + 1}: {rowText}");
                    lines.Add($"  Parsed: {record.Date} | {record.Type} | {record.Aircraft} | {record.DepartureTime} | {record.From} -> {record.To} | {record.Duration}");
                }
                else
                {
                    lines.Add($"Row {index + 1}: {rowText}");
                    lines.Add($"  Rejected: {GetRejectReason(tokens)}");
                }
            }

            return string.Join("\n", lines);
        }

        private string GetRejectReason(List<string> tokens)
        {
            var normalizedTokens = tokens.Select(NormalizeText).ToList();
            if (normalizedTokens.FirstOrDefault(IsDateToken) == null)
                return "Missing date";

            var timeTokens = normalizedTokens.Where(IsTimeToken).ToList();
            if (timeTokens.Count < 2)
                return "Missing departure or arrival time";

            if (normalizedTokens.Count(IsIcaoCode) < 2)
                return "Missing from/to ICAO codes";

            return "Could not build a complete flight record";
        }

        private List<List<string>> GetRowTokenGroups(OcrResult result)
        {
            var words = result.Lines
                .SelectMany(l => l.Words)
                .Select(w => new WordBox
                {
                    Text = NormalizeText(w.Text),
                    X = w.BoundingRect.X,
                    Y = w.BoundingRect.Y
                })
                .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                .OrderBy(w => w.Y)
                .ThenBy(w => w.X)
                .ToList();

            var rows = new List<List<WordBox>>();
            const double threshold = 12.0;

            foreach (var word in words)
            {
                var row = rows
                    .OrderBy(r => Math.Abs(r.Average(w => w.Y) - word.Y))
                    .FirstOrDefault(r => Math.Abs(r.Average(w => w.Y) - word.Y) < threshold);

                if (row == null)
                {
                    rows.Add(new List<WordBox> { word });
                }
                else
                {
                    row.Add(word);
                }
            }

            foreach (var row in rows)
            {
                row.Sort((a, b) => a.X.CompareTo(b.X));
            }

            return rows
                .Select(row => row.Select(w => w.Text ?? string.Empty).Where(t => !string.IsNullOrWhiteSpace(t)).ToList())
                .Where(tokens => tokens.Count > 0)
                .ToList();
        }

        // ✈️ Parse a single row into a flight record
        private FlightRecord? ParseRow(List<string> tokens)
        {
            var normalizedTokens = tokens
                .Select(NormalizeText)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList();

            string? date = normalizedTokens.FirstOrDefault(IsDateToken);
            if (date == null)
                return null;

            int dateIndex = normalizedTokens.IndexOf(date);
            int searchStart = dateIndex + 1;
            if (searchStart >= normalizedTokens.Count)
                return null;

            int typeIndex = normalizedTokens.FindIndex(searchStart, t =>
                t.Contains("AIRCRAFT") || t.Contains("HELICOPTER") || t.Contains("PLANE") || t.Contains("HELI") || t.Contains("AIR"));

            string type = typeIndex >= 0 ? normalizedTokens[typeIndex] : string.Empty;
            int aircraftStart = typeIndex >= 0 ? typeIndex + 1 : searchStart;

            var timeTokens = normalizedTokens.Where(IsTimeToken).ToList();
            string? departureTime = timeTokens.ElementAtOrDefault(0);
            string? arrivalTime = timeTokens.ElementAtOrDefault(1);
            string? duration = normalizedTokens.FirstOrDefault(IsDurationToken) ?? string.Empty;

            var icaoCodes = normalizedTokens.Where(IsIcaoCode).ToList();
            string? from = icaoCodes.ElementAtOrDefault(0);
            string? to = icaoCodes.ElementAtOrDefault(1);

            int aircraftEnd = normalizedTokens.FindIndex(aircraftStart, t => IsTimeToken(t) || IsIcaoCode(t) || IsDurationToken(t));
            if (aircraftEnd < 0)
                aircraftEnd = normalizedTokens.Count;

            var aircraftTokens = normalizedTokens
                .Skip(aircraftStart)
                .Take(aircraftEnd - aircraftStart)
                .Where(t => t != "|")
                .ToList();

            string aircraft = string.Join(" ", aircraftTokens).Trim();
            if (string.IsNullOrWhiteSpace(aircraft) && aircraftTokens.Count == 0 && normalizedTokens.Count > aircraftStart)
            {
                aircraft = normalizedTokens[aircraftStart];
            }

            if (string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(to) || string.IsNullOrWhiteSpace(departureTime) || string.IsNullOrWhiteSpace(arrivalTime))
                return null;

            return new FlightRecord
            {
                Date = date,
                Type = type,
                Aircraft = aircraft,
                DepartureTime = departureTime,
                From = from,
                ArrivalTime = arrivalTime,
                To = to,
                Duration = duration
            };
        }

        private static readonly string[] DatePatterns = new[]
        {
            "yyyy-MM-dd",
            "dd.MM.yyyy",
            "dd/MM/yyyy",
            "MM/dd/yyyy",
            "dd-MM-yyyy",
            "yyyyMMdd"
        };

        private static string NormalizeText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var normalized = text.Trim().ToUpperInvariant();
            normalized = normalized.Replace("|", string.Empty);
            return normalized;
        }

        private static bool IsDateToken(string token)
        {
            if (DateTime.TryParse(token, out _))
                return true;

            return DatePatterns.Any(pattern => DateTime.TryParseExact(token, pattern, null, System.Globalization.DateTimeStyles.None, out _));
        }

        private static bool IsTimeToken(string token)
        {
            return Regex.IsMatch(token, @"^\d{1,2}[:\.]\d{2}$");
        }

        private static bool IsDurationToken(string token)
        {
            return Regex.IsMatch(token, @"^\d{1,2}[:\.]\d{2}[:\.]\d{2}$");
        }

        private static bool IsIcaoCode(string token)
        {
            return Regex.IsMatch(token, @"^[A-Z]{4}$");
        }

        // Export flights to JSON or XML
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (_flightRecords.Count == 0)
            {
                MessageBox.Show("No flight records to export.");
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = FormatComboBox.SelectedIndex == 0 ? "JSON files (*.json)|*.json" : "XML files (*.xml)|*.xml",
                DefaultExt = FormatComboBox.SelectedIndex == 0 ? "json" : "xml"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    if (FormatComboBox.SelectedIndex == 0) // JSON
                    {
                        var json = JsonSerializer.Serialize(_flightRecords, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(saveFileDialog.FileName, json);
                    }
                    else // XML
                    {
                        var serializer = new XmlSerializer(typeof(List<FlightRecord>));
                        using var writer = new StreamWriter(saveFileDialog.FileName);
                        serializer.Serialize(writer, _flightRecords);
                    }
                    MessageBox.Show("Flights exported successfully.");
                    MyImageControl.Source = null;
                    _flightRecords.Clear();
                    UpdateRouteCount();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting: {ex.Message}");
                }
            }
        }

        // Load existing file and append current records
        private void LoadAppend_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "JSON or XML files (*.json;*.xml)|*.json;*.xml"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    List<FlightRecord> existingRecords;
                    if (Path.GetExtension(openFileDialog.FileName).ToLower() == ".json")
                    {
                        var json = File.ReadAllText(openFileDialog.FileName);
                        existingRecords = JsonSerializer.Deserialize<List<FlightRecord>>(json) ?? new List<FlightRecord>();
                    }
                    else
                    {
                        var serializer = new XmlSerializer(typeof(List<FlightRecord>));
                        using var reader = new StreamReader(openFileDialog.FileName);
                        existingRecords = (List<FlightRecord>?)serializer.Deserialize(reader) ?? new List<FlightRecord>();
                    }

                    var newRecords = _flightRecords.Except(existingRecords).ToList();
                    if (newRecords.Count == 0)
                    {
                        MessageBox.Show("No new records to append. All current records are already in the file.");
                        return;
                    }

                    existingRecords.AddRange(newRecords);
                    _flightRecords = existingRecords;

                    // Save back
                    if (Path.GetExtension(openFileDialog.FileName).ToLower() == ".json")
                    {
                        var json = JsonSerializer.Serialize(_flightRecords, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(openFileDialog.FileName, json);
                    }
                    else
                    {
                        var serializer = new XmlSerializer(typeof(List<FlightRecord>));
                        using var writer = new StreamWriter(openFileDialog.FileName);
                        serializer.Serialize(writer, _flightRecords);
                    }

                    MessageBox.Show($"{newRecords.Count} new records appended successfully.");
                    MyImageControl.Source = null;
                    _flightRecords.Clear();
                    UpdateRouteCount();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading/appending: {ex.Message}");
                }
            }
        }
    }

    // 📦 Helper class for OCR word positions
    public class WordBox
    {
        public string? Text { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
    }

    // ✈️ Flight model
    [Serializable]
    public class FlightRecord : IEquatable<FlightRecord>
    {
        public string? Date { get; set; }
        public string? Type { get; set; }
        public string? Aircraft { get; set; }
        public string? DepartureTime { get; set; }
        public string? From { get; set; }
        public string? ArrivalTime { get; set; }
        public string? To { get; set; }
        public string? Duration { get; set; }

        public bool Equals(FlightRecord? other)
        {
            if (other is null) return false;
            return string.Equals(Date, other.Date, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(DepartureTime, other.DepartureTime, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(From, other.From, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(To, other.To, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj) => Equals(obj as FlightRecord);

        public override int GetHashCode() => HashCode.Combine(Date?.ToLowerInvariant(), DepartureTime?.ToLowerInvariant(), From?.ToLowerInvariant(), To?.ToLowerInvariant());
    }
}