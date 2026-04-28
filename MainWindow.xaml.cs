using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
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

            if (DebugModeCheckBox.IsChecked == true)
            {
                OutputTextBox.Text = GetDebugOutput(result);
                return;
            }

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
            // Preprocess before first await so it runs on the UI thread
            var preprocessed = PreprocessForOcr(bitmap);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(preprocessed));

            using var stream = new MemoryStream();
            encoder.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);

            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream.AsRandomAccessStream());
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync();

            var ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
            return await ocrEngine.RecognizeAsync(softwareBitmap);
        }

        // Scale up 2× and invert colors so Windows OCR gets dark text on a light background
        private static BitmapSource PreprocessForOcr(BitmapSource source)
        {
            const double scale = 2.0;
            int w = (int)(source.PixelWidth * scale);
            int h = (int)(source.PixelHeight * scale);

            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
                dc.DrawImage(source, new Rect(0, 0, w, h));

            var scaled = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
            scaled.Render(visual);

            // Convert to Bgra32 and invert R, G, B channels
            var converted = new FormatConvertedBitmap(scaled, PixelFormats.Bgra32, null, 0);
            int stride = w * 4;
            byte[] pixels = new byte[h * stride];
            converted.CopyPixels(pixels, stride, 0);

            for (int i = 0; i < pixels.Length; i += 4)
            {
                pixels[i]     = (byte)(255 - pixels[i]);     // B
                pixels[i + 1] = (byte)(255 - pixels[i + 1]); // G
                pixels[i + 2] = (byte)(255 - pixels[i + 2]); // R
                // alpha unchanged
            }

            return BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, pixels, stride);
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
            const double threshold = 20.0;

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

            var rowTokens = rows
                .Select(row => row.Select(w => w.Text ?? string.Empty).Where(t => !string.IsNullOrWhiteSpace(t)).ToList())
                .Where(tokens => tokens.Count > 0)
                .ToList();

            rowTokens = MergeSplitRows(rowTokens);

            var records = new List<FlightRecord>();

            foreach (var tokens in rowTokens)
            {
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

        private string GetDebugOutput(OcrResult result)
        {
            var sb = new StringBuilder();

            // 1. Raw OCR words with positions
            sb.AppendLine("=== RAW OCR WORDS (x, y, text) ===");
            foreach (var line in result.Lines)
                foreach (var word in line.Words)
                    sb.AppendLine($"  ({word.BoundingRect.X:F0}, {word.BoundingRect.Y:F0})  \"{word.Text}\"  → normalized: \"{NormalizeText(word.Text)}\"");

            // 2. Y-grouped rows before merge
            var words = result.Lines
                .SelectMany(l => l.Words)
                .Select(w => new WordBox { Text = NormalizeText(w.Text), X = w.BoundingRect.X, Y = w.BoundingRect.Y })
                .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                .OrderBy(w => w.Y).ThenBy(w => w.X)
                .ToList();

            var rawRows = new List<List<WordBox>>();
            const double threshold = 20.0;
            foreach (var word in words)
            {
                var row = rawRows
                    .OrderBy(r => Math.Abs(r.Average(w => w.Y) - word.Y))
                    .FirstOrDefault(r => Math.Abs(r.Average(w => w.Y) - word.Y) < threshold);
                if (row == null) rawRows.Add(new List<WordBox> { word });
                else row.Add(word);
            }
            foreach (var row in rawRows) row.Sort((a, b) => a.X.CompareTo(b.X));

            sb.AppendLine("\n=== Y-GROUPED ROWS (before merge) ===");
            for (int i = 0; i < rawRows.Count; i++)
            {
                var avgY = rawRows[i].Average(w => w.Y);
                var tokens = rawRows[i].Select(w => w.Text ?? "").ToList();
                sb.AppendLine($"  Row {i + 1} (avgY={avgY:F0}): {string.Join(" | ", tokens)}");
            }

            // 3. Rows after merge
            var tokenRows = rawRows
                .Select(row => row.Select(w => w.Text ?? "").Where(t => !string.IsNullOrWhiteSpace(t)).ToList())
                .Where(t => t.Count > 0)
                .ToList();
            var merged = MergeSplitRows(tokenRows);

            sb.AppendLine("\n=== ROWS AFTER MERGE ===");
            for (int i = 0; i < merged.Count; i++)
                sb.AppendLine($"  Row {i + 1}: {string.Join(" | ", merged[i])}");

            // 4. Parse result per merged row
            sb.AppendLine("\n=== PARSE RESULTS ===");
            for (int i = 0; i < merged.Count; i++)
            {
                var tokens = merged[i];
                var record = ParseRow(tokens);
                if (record != null)
                    sb.AppendLine($"  Row {i + 1}: OK  →  {record.Date} | {record.Aircraft} | {record.From} → {record.To} | {record.Duration}");
                else
                    sb.AppendLine($"  Row {i + 1}: REJECTED  ({GetRejectReason(tokens)})  tokens: {string.Join(" | ", tokens)}");
            }

            return sb.ToString();
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

            var airportCount = normalizedTokens
                .Select(ExtractIcaoCode)
                .Count(code => !string.IsNullOrWhiteSpace(code));
            if (airportCount == 0)
                return "No ICAO airport code found";
            if (airportCount == 1)
                return "Only one airport code found (missing To)";

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
            const double threshold = 20.0;

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

        private static List<List<string>> MergeSplitRows(List<List<string>> rows)
        {
            var result = new List<List<string>>();
            int i = 0;

            while (i < rows.Count)
            {
                var current = rows[i];
                while (i + 1 < rows.Count && ShouldMergeRows(current, rows[i + 1]))
                {
                    current = RemoveDuplicateAdjacentTokens(current.Concat(rows[i + 1]).ToList());
                    i++;
                }
                result.Add(current);
                i++;
            }

            return result;
        }

        private static bool ShouldMergeRows(List<string> left, List<string> right)
        {
            if (left.Count == 0 || right.Count == 0)
                return false;

            var leftNorm = left.Select(NormalizeText).ToList();
            var rightNorm = right.Select(NormalizeText).ToList();

            if (StartsWithDate(leftNorm) && !StartsWithDate(rightNorm))
            {
                return IsContinuationRow(leftNorm, rightNorm);
            }

            if (!StartsWithDate(leftNorm) && StartsWithDate(rightNorm))
            {
                return IsContinuationRow(rightNorm, leftNorm);
            }

            if (!leftNorm[0].Equals(rightNorm[0], StringComparison.OrdinalIgnoreCase))
                return false;

            if (leftNorm.Count > 1 && rightNorm.Count > 1 && !leftNorm[1].Equals(rightNorm[1], StringComparison.OrdinalIgnoreCase))
                return false;

            int matchingPrefix = leftNorm.Zip(rightNorm, (a, b) => a == b ? 1 : 0).Take(6).Sum();
            if (matchingPrefix < 4)
                return false;

            return IsPrefix(leftNorm, rightNorm) || IsPrefix(rightNorm, leftNorm);
        }

        private static bool IsContinuationRow(List<string> rowWithDate, List<string> possibleContinuation)
        {
            if (possibleContinuation.Count == 0)
                return false;

            if (!HasOnlyFromCode(rowWithDate))
                return false;

            if (StartsWithDate(possibleContinuation))
                return false;

            return IsTimeToken(possibleContinuation[0]) || ExtractIcaoCode(possibleContinuation[0]) != null;
        }

        // Returns true when a row has exactly one extractable airport code — meaning it still needs the "to" airport
        private static bool HasOnlyFromCode(List<string> tokens)
        {
            var normalized = tokens.Select(NormalizeText).ToList();
            int firstAirportIndex = normalized.FindIndex(t => ExtractIcaoCode(t) != null);
            if (firstAirportIndex < 0)
                return false;

            var airportCodes = normalized
                .Skip(firstAirportIndex)
                .Select(ExtractIcaoCode)
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .ToList();

            return airportCodes.Count == 1;
        }

        private static bool IsPrefix(List<string> a, List<string> b)
        {
            if (a.Count >= b.Count)
                return false;

            for (int i = 0; i < a.Count; i++)
            {
                if (!a[i].Equals(b[i], StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }

        private static bool StartsWithDate(List<string> tokens)
        {
            return tokens.Count > 0 && IsDateToken(tokens[0]);
        }

        private static List<string> RemoveDuplicateAdjacentTokens(List<string> tokens)
        {
            var result = new List<string>();

            foreach (var token in tokens)
            {
                if (result.Count == 0 || !result.Last().Equals(token, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(token);
                }
            }

            return result;
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

            // Determine the aircraft name boundary using whichever anchor comes first:
            //   • A time token (departure time) — the normal case
            //   • A bracketed airport code [XXXX] — fallback when departure time is absent
            // Plain ICAO codes (e.g. "A320") are NOT used as anchors to avoid false positives
            // with aircraft model numbers.
            int firstTimeIndex = normalizedTokens.FindIndex(aircraftStart, IsTimeToken);
            int firstBracketedIndex = normalizedTokens.FindIndex(aircraftStart, IsBracketedIcaoCode);

            int aircraftEnd;
            string? departureTime;
            int airportSectionStart;

            if (firstTimeIndex >= 0 && (firstBracketedIndex < 0 || firstTimeIndex < firstBracketedIndex))
            {
                // Time comes first (or no bracketed code): departure time found
                departureTime = normalizedTokens[firstTimeIndex];
                aircraftEnd = firstTimeIndex;
                airportSectionStart = firstTimeIndex + 1;
            }
            else if (firstBracketedIndex >= 0)
            {
                // Bracketed airport code precedes any time token: no departure time logged
                departureTime = null;
                aircraftEnd = firstBracketedIndex;
                airportSectionStart = firstBracketedIndex;
            }
            else
            {
                return null;
            }

            // Aircraft name — strip single-char tokens (OCR misreads of the '|' column separator)
            var aircraftTokens = normalizedTokens
                .Skip(aircraftStart)
                .Take(aircraftEnd - aircraftStart)
                .Where(t => t.Length > 1)
                .ToList();
            string aircraft = string.Join(" ", aircraftTokens).Trim();

            // Collect all ICAO codes in the airport section
            var airportCodes = normalizedTokens
                .Skip(airportSectionStart)
                .Select(ExtractIcaoCode)
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .ToList();

            string? from = airportCodes.ElementAtOrDefault(0);
            string? to = airportCodes.ElementAtOrDefault(1);

            if (string.IsNullOrWhiteSpace(from))
                return null;

            // Arrival time: first time token after the from-airport position
            int firstAirportPos = normalizedTokens.FindIndex(airportSectionStart, t => ExtractIcaoCode(t) != null);
            int arrivalIndex = firstAirportPos >= 0 ? normalizedTokens.FindIndex(firstAirportPos + 1, IsTimeToken) : -1;
            string? arrivalTime = arrivalIndex >= 0 ? normalizedTokens[arrivalIndex] : null;

            string? duration = normalizedTokens.FirstOrDefault(IsDurationToken) ?? string.Empty;

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

        private static bool IsBracketedIcaoCode(string token)
        {
            return Regex.IsMatch(token, @"^\[[A-Z0-9]{3,5}\]$");
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

// Common OCR misreads where two characters are substituted for one
        private static readonly (string From, string To)[] OcrDigraphFixes =
        {
            ("IJ", "U"),  // "U" read as "IJ" (e.g. KBIJR → KBUR)
            ("VV", "W"),
            ("RN", "M"),
        };

        private static string? ExtractIcaoCode(string? token)
        {
            if (string.IsNullOrWhiteSpace(token))
                return null;

            // Standard 3-4 char code (ICAO, IATA, FAA)
            if (Regex.IsMatch(token, @"^(?=.*[A-Z])[A-Z0-9]{3,4}$"))
                return token;

            // Bracketed format like [EFSI] Vicinity or [CEK9B] (up to 5 chars)
            var match = Regex.Match(token, @"\[([A-Z0-9]{3,5})\]");
            if (match.Success)
                return match.Groups[1].Value;

            // 5-char alphanumeric with at least one letter: try OCR digraph correction first
            // (e.g., "KBIJR" → "KBUR" via IJ→U), then accept as a legitimate 5-char identifier
            // (Transport Canada, some FAA private/small airfield codes)
            if (Regex.IsMatch(token, @"^(?=.*[A-Z])[A-Z0-9]{5}$"))
            {
                foreach (var (from, to) in OcrDigraphFixes)
                {
                    var corrected = token.Replace(from, to);
                    if (Regex.IsMatch(corrected, @"^(?=.*[A-Z])[A-Z0-9]{3,4}$"))
                        return corrected;
                }
                return token;
            }

            return null;
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

        // Convert format between JSON and XML
        private void ConvertFormat_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "JSON or XML files (*.json;*.xml)|*.json;*.xml"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    List<FlightRecord> records;
                    string sourceFormat = Path.GetExtension(openFileDialog.FileName).ToLower();

                    // Load the file
                    if (sourceFormat == ".json")
                    {
                        var json = File.ReadAllText(openFileDialog.FileName);
                        records = JsonSerializer.Deserialize<List<FlightRecord>>(json) ?? new List<FlightRecord>();
                    }
                    else
                    {
                        var serializer = new XmlSerializer(typeof(List<FlightRecord>));
                        using var reader = new StreamReader(openFileDialog.FileName);
                        records = (List<FlightRecord>?)serializer.Deserialize(reader) ?? new List<FlightRecord>();
                    }

                    if (records.Count == 0)
                    {
                        MessageBox.Show("No flight records found in the file.");
                        return;
                    }

                    // Ask for target format
                    string targetFormat = sourceFormat == ".json" ? "XML" : "JSON";
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = targetFormat == "JSON" ? "JSON files (*.json)|*.json" : "XML files (*.xml)|*.xml",
                        DefaultExt = targetFormat == "JSON" ? "json" : "xml",
                        FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_converted"
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        // Save in target format
                        if (targetFormat == "JSON")
                        {
                            var json = JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });
                            File.WriteAllText(saveFileDialog.FileName, json);
                        }
                        else
                        {
                            var serializer = new XmlSerializer(typeof(List<FlightRecord>));
                            using var writer = new StreamWriter(saveFileDialog.FileName);
                            serializer.Serialize(writer, records);
                        }

                        MessageBox.Show($"File converted and saved as {targetFormat} successfully.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error converting file: {ex.Message}");
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