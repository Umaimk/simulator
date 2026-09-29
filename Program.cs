// ============================================================================
// Program.cs — NICVD Cardiac OPD Single-Server Queueing Analysis
//
// SELECTED QUEUEING MODEL: **M/G/1** (single server)
//   M = arrivals are Poisson / Markovian        -> confirmed by chi-square GOF
//   G = service times are a General distribution -> service is NOT exponential
//        (sample coefficient of variation is far below 1)
//   1 = a single server (one doctor). Note: the standard queueing-theory
//       terms lambda/mu/L/Lq/W/Wq are defined generically over "customers",
//       but this program's console output labels them "patients" since
//       that's what they represent in this OPD dataset.
// M/M/1 is also printed, but only as a reference comparison — it is not the
// model the data supports, because service time fails the exponential check.
//
// THIS PROGRAM NOW IMPORTS ALL THREE SOURCE FILES DIRECTLY:
//   1. OPD-Data-Group1.xlsx        -> read with ClosedXML (the raw dataset)
//   2. Chi_Square_Analysison...pdf -> read with PdfPig, chi-square figures
//                                      parsed out of the extracted text
//   3. OPD-SPSS.pdf                -> read with PdfPig, used to CROSS-CHECK
//                                      the service-time mean/SD computed
//                                      directly from the Excel data
//
// If a PDF's text cannot be reliably parsed (PDF table layouts can extract
// in an unpredictable order), the program prints a clear WARNING and falls
// back to the exact figures transcribed from that PDF, so the analysis
// never silently uses wrong numbers.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace OpdQueueing
{
    public class PatientRecord
    {
        public int SerialNo;
        public string Name = string.Empty;
        public double ArrivalMin;
        public double ServiceStartMin;
        public double ServiceEndMin;

        public double ServiceTime => ServiceEndMin - ServiceStartMin;
        public double QueueWaitTime => ServiceStartMin - ArrivalMin;
        public double SystemTime => ServiceEndMin - ArrivalMin;
    }

    public class ChiSquareResult
    {
        public string TestName = string.Empty;
        public double ChiSquareStat;
        public int DegreesOfFreedom;
        public double CriticalValue;
        public double Alpha = 0.05;
        public bool RejectNull => ChiSquareStat > CriticalValue;
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            // ---- Resolve input file paths (override via command-line args) ----
            string excelPath = args.Length > 0 ? args[0] : "OPD-Data-Group1.xlsx";
            string chiSquarePdfPath = args.Length > 1 ? args[1] : "Chi_Square_Analysison_arrival_time.pdf";
            string spssPdfPath = args.Length > 2 ? args[2] : "OPD-SPSS.pdf";

            Console.WriteLine("================================================================");
            Console.WriteLine(" NICVD CARDIAC OPD - SINGLE-SERVER QUEUEING ANALYSIS");
            Console.WriteLine(" SELECTED MODEL: M/G/1  (Poisson arrivals, General service, 1 server)");
            Console.WriteLine("================================================================\n");

            // ---------------------------------------------------------------
            // STEP 1: IMPORT the raw dataset from the actual Excel file
            // ---------------------------------------------------------------
            List<PatientRecord> patients;
            try
            {
                patients = LoadDatasetFromExcel(excelPath);
                Console.WriteLine($"[OK] Loaded {patients.Count} patient records from '{excelPath}'.\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FATAL] Could not read Excel file '{excelPath}': {ex.Message}");
                return;
            }

            // ---------------------------------------------------------------
            // STEP 2: Derived per-patient quantities (computed, not imported)
            // ---------------------------------------------------------------
            var interArrivalTimes = new List<double>();
            for (int i = 1; i < patients.Count; i++)
                interArrivalTimes.Add(patients[i].ArrivalMin - patients[i - 1].ArrivalMin);

            var serviceTimes = patients.Select(p => p.ServiceTime).ToList();
            var queueWaits = patients.Select(p => p.QueueWaitTime).ToList();
            var systemTimes = patients.Select(p => p.SystemTime).ToList();

            double meanInterArrival = interArrivalTimes.Average();
            double meanServiceTime = serviceTimes.Average();
            double sdServiceTime = StdDev(serviceTimes);
            double meanQueueWait = queueWaits.Average();
            double meanSystemTime = systemTimes.Average();

            Console.WriteLine("--- Derived sample statistics (computed from the imported Excel data) ---");
            Console.WriteLine($"N (patients)                  = {patients.Count}");
            Console.WriteLine($"Mean inter-arrival time       = {meanInterArrival:F4} min");
            Console.WriteLine($"Mean service time              = {meanServiceTime:F4} min");
            Console.WriteLine($"SD of service time              = {sdServiceTime:F4} min");
            Console.WriteLine($"Coefficient of variation (Cs) = {sdServiceTime / meanServiceTime:F4}  " +
                               "(Exponential needs Cs = 1 -> service is NOT exponential -> supports 'G', not 'M')");
            Console.WriteLine($"Mean observed queue wait       = {meanQueueWait:F4} min");
            Console.WriteLine($"Mean observed system time      = {meanSystemTime:F4} min\n");

            // ---------------------------------------------------------------
            // STEP 3: IMPORT the SPSS PDF and cross-check service time stats
            // ---------------------------------------------------------------
            var (spssMean, spssSd, spssParsed) = TryLoadServiceTimeDescriptivesFromSpss(spssPdfPath);
            Console.WriteLine("--- Cross-check against SPSS PDF descriptives ---");
            if (spssParsed)
            {
                Console.WriteLine($"[OK] Parsed from '{spssPdfPath}': Service Time mean={spssMean:F2}, SD={spssSd:F3}");
                bool meanMatch = Math.Abs(spssMean - meanServiceTime) < 0.05;
                bool sdMatch = Math.Abs(spssSd - sdServiceTime) < 0.05;
                Console.WriteLine(meanMatch ? "  Mean matches Excel-derived value: PASS" : "  Mean DOES NOT match Excel-derived value: CHECK DATA");
                Console.WriteLine(sdMatch ? "  SD matches Excel-derived value: PASS" : "  SD DOES NOT match Excel-derived value: CHECK DATA");
            }
            else
            {
                Console.WriteLine($"[WARNING] Could not reliably parse Service Time row out of '{spssPdfPath}'.");
                Console.WriteLine("          Proceeding with the value computed directly from the Excel data.");
            }
            Console.WriteLine();

            // ---------------------------------------------------------------
            // STEP 4: Arrival and service rates
            // ---------------------------------------------------------------
            double lambda = 1.0 / meanInterArrival;
            double mu = 1.0 / meanServiceTime;

            Console.WriteLine("--- Arrival and service rates ---");
            Console.WriteLine($"lambda = 1 / mean inter-arrival = {lambda:F4} patients/min ({lambda * 60:F3} /hour)");
            Console.WriteLine($"mu     = 1 / mean service time  = {mu:F4} patients/min ({mu * 60:F3} /hour)\n");

            // ---------------------------------------------------------------
            // STEP 5: Utilization
            // ---------------------------------------------------------------
            double rho = lambda / mu;
            Console.WriteLine("--- Utilization ---");
            Console.WriteLine($"rho = lambda / mu = {rho:F4}");
            Console.WriteLine(rho < 1 ? "System is STABLE (rho < 1).\n"
                                       : "System is UNSTABLE (rho >= 1): steady-state formulas below do not apply.\n");

            // ---------------------------------------------------------------
            // STEP 6: Theoretical M/M/1 (reference comparison only)
            // ---------------------------------------------------------------
            double L_mm1 = double.NaN, Lq_mm1 = double.NaN, W_mm1 = double.NaN, Wq_mm1 = double.NaN;
            if (rho < 1)
            {
                L_mm1 = rho / (1 - rho);
                Lq_mm1 = (rho * rho) / (1 - rho);
                W_mm1 = 1.0 / (mu - lambda);
                Wq_mm1 = rho / (mu - lambda);
            }
            Console.WriteLine("--- Theoretical M/M/1 measures (REFERENCE ONLY - not the selected model) ---");
            Console.WriteLine($"L={L_mm1:F4} patients, Lq={Lq_mm1:F4} patients, W={W_mm1:F4} min, Wq={Wq_mm1:F4} min\n");

            // ---------------------------------------------------------------
            // STEP 7: Theoretical M/G/1 (Pollaczek-Khinchine) — SELECTED MODEL
            // ---------------------------------------------------------------
            double varServiceTime = sdServiceTime * sdServiceTime;
            double Lq_mg1 = double.NaN, Wq_mg1 = double.NaN, W_mg1 = double.NaN, L_mg1 = double.NaN;
            if (rho < 1)
            {
                Lq_mg1 = (lambda * lambda * varServiceTime + rho * rho) / (2 * (1 - rho));
                Wq_mg1 = Lq_mg1 / lambda;
                W_mg1 = Wq_mg1 + (1.0 / mu);
                L_mg1 = lambda * W_mg1;
            }
            Console.WriteLine("--- Theoretical M/G/1 (Pollaczek-Khinchine) measures — SELECTED MODEL ---");
            Console.WriteLine($"Lq = (lambda^2*sigma_s^2 + rho^2) / (2*(1-rho)) = {Lq_mg1:F4} patients");
            Console.WriteLine($"Wq = Lq / lambda                                = {Wq_mg1:F4} min");
            Console.WriteLine($"W  = Wq + 1/mu                                  = {W_mg1:F4} min");
            Console.WriteLine($"L  = lambda * W                                 = {L_mg1:F4} patients\n");

            // ---------------------------------------------------------------
            // STEP 8: Empirical measures directly from imported data
            // ---------------------------------------------------------------
            double windowStart = patients.Min(p => p.ArrivalMin);
            double windowEnd = patients.Max(p => p.ServiceEndMin);
            double totalWindow = windowEnd - windowStart;

            double L_obs = TimeAverageCount(patients.Select(p => (p.ArrivalMin, p.ServiceEndMin)).ToList(), windowStart, windowEnd);
            double Lq_obs = TimeAverageCount(patients.Select(p => (p.ArrivalMin, p.ServiceStartMin)).ToList(), windowStart, windowEnd);
            double lambdaObsWindow = patients.Count / totalWindow;

            Console.WriteLine("--- Empirical (observed) measures, from imported data, no distribution assumed ---");
            Console.WriteLine($"Observation window = {windowStart:F0} to {windowEnd:F0} min ({totalWindow:F0} min span)");
            Console.WriteLine($"L_obs  = {L_obs:F4} patients   Lq_obs = {Lq_obs:F4} patients");
            Console.WriteLine($"W_obs  = {meanSystemTime:F4} min   Wq_obs = {meanQueueWait:F4} min");
            Console.WriteLine("NOTE: much larger than theoretical values -- service began well after the first");
            Console.WriteLine("patients arrived, so a startup backlog formed; this session was not steady-state.\n");

            // ---------------------------------------------------------------
            // STEP 9: Little's Law validation (empirical)
            // ---------------------------------------------------------------
            Console.WriteLine("--- Little's Law validation (empirical) ---");
            Console.WriteLine($"lambda_obs (N / window span) = {lambdaObsWindow:F4} patients/min");
            Console.WriteLine($"lambda_obs * W_obs  = {lambdaObsWindow * meanSystemTime:F4}  (vs L_obs  = {L_obs:F4})");
            Console.WriteLine($"lambda_obs * Wq_obs = {lambdaObsWindow * meanQueueWait:F4}  (vs Lq_obs = {Lq_obs:F4})\n");

            // ---------------------------------------------------------------
            // STEP 10: IMPORT the Chi-Square PDF results
            // ---------------------------------------------------------------
            var (poissonTest, expTest, chiParsed) = TryLoadChiSquareResultsFromPdf(chiSquarePdfPath);
            Console.WriteLine("--- Distribution test summary (imported from Chi-Square PDF) ---");
            Console.WriteLine(chiParsed
                ? $"[OK] Parsed chi-square figures from '{chiSquarePdfPath}'."
                : $"[WARNING] Could not reliably parse '{chiSquarePdfPath}'; using transcribed fallback values.");
            foreach (var t in new[] { poissonTest, expTest })
            {
                Console.WriteLine($"{t.TestName}: chi2={t.ChiSquareStat}, df={t.DegreesOfFreedom}, " +
                                   $"critical={t.CriticalValue}, alpha={t.Alpha} => " +
                                   (t.RejectNull ? "Reject H0" : "Fail to reject H0"));
            }
            Console.WriteLine("Service time GOF: NOT AVAILABLE - no chi-square test was supplied for this variable");
            Console.WriteLine("(this is exactly why the service side of the model is treated as 'G', not 'M').\n");

            // ---------------------------------------------------------------
            // STEP 11: Final results table
            // ---------------------------------------------------------------
            Console.WriteLine("================================================================");
            Console.WriteLine(" FINAL RESULTS TABLE  |  SELECTED QUEUEING MODEL: M/G/1 (single server)");
            Console.WriteLine("================================================================");
            Console.WriteLine($"{"Measure",-12}{"M/M/1(ref)",-14}{"M/G/1(SELECTED)",-18}{"Empirical",-12}Unit");
            Console.WriteLine($"{"lambda",-12}{lambda,-14:F4}{lambda,-18:F4}{lambdaObsWindow,-12:F4}patients/min");
            Console.WriteLine($"{"mu",-12}{mu,-14:F4}{mu,-18:F4}{"-",-12}patients/min");
            Console.WriteLine($"{"rho",-12}{rho,-14:F4}{rho,-18:F4}{"-",-12}(none)");
            Console.WriteLine($"{"L",-12}{L_mm1,-14:F4}{L_mg1,-18:F4}{L_obs,-12:F4}patients");
            Console.WriteLine($"{"Lq",-12}{Lq_mm1,-14:F4}{Lq_mg1,-18:F4}{Lq_obs,-12:F4}patients");
            Console.WriteLine($"{"W",-12}{W_mm1,-14:F4}{W_mg1,-18:F4}{meanSystemTime,-12:F4}minutes");
            Console.WriteLine($"{"Wq",-12}{Wq_mm1,-14:F4}{Wq_mg1,-18:F4}{meanQueueWait,-12:F4}minutes");
        }

        // ====================================================================
        // IMPORT #1: Excel dataset
        // ====================================================================
        private static List<PatientRecord> LoadDatasetFromExcel(string path)
        {
            const int originMinutes = 7 * 60; // times are reported relative to 7:00 AM

            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheet(1);
            var usedRows = ws.RangeUsed()!.RowsUsed();

            var list = new List<PatientRecord>();
            foreach (var row in usedRows)
            {
                var snCell = row.Cell(1);
                if (!snCell.TryGetValue<int>(out int sn)) continue; // skips title/header rows

                string name = row.Cell(2).GetString();
                TimeSpan arrival = row.Cell(3).GetDateTime().TimeOfDay;
                TimeSpan start = row.Cell(4).GetDateTime().TimeOfDay;
                TimeSpan end = row.Cell(5).GetDateTime().TimeOfDay;

                list.Add(new PatientRecord
                {
                    SerialNo = sn,
                    Name = name,
                    ArrivalMin = arrival.TotalMinutes - originMinutes,
                    ServiceStartMin = start.TotalMinutes - originMinutes,
                    ServiceEndMin = end.TotalMinutes - originMinutes
                });
            }
            return list.OrderBy(p => p.SerialNo).ToList();
        }

        // ====================================================================
        // IMPORT #2: Chi-Square PDF -> parsed statistics (with safe fallback)
        // ====================================================================
        private static (ChiSquareResult poisson, ChiSquareResult exponential, bool parsedOk)
            TryLoadChiSquareResultsFromPdf(string pdfPath)
        {
            // Fallback values transcribed directly from the supplied PDF,
            // used only if text extraction/parsing does not succeed.
            var poisson = new ChiSquareResult
            {
                TestName = "Poisson GOF (arrivals per interval)",
                ChiSquareStat = 5.0859,
                DegreesOfFreedom = 6,
                CriticalValue = 12.5916
            };
            var exp = new ChiSquareResult
            {
                TestName = "Exponential GOF (inter-arrival times)",
                ChiSquareStat = 1.5641,
                DegreesOfFreedom = 1,
                CriticalValue = 3.8415
            };

            string text;
            try
            {
                text = ExtractPdfText(pdfPath);
            }
            catch
            {
                return (poisson, exp, false);
            }

            try
            {
                var calculated = Regex.Matches(text, @"Calculated(?:\s+Chi-Square)?\s+([\d.]+)");
                var critical = Regex.Matches(text, @"Critical value\s+([\d.]+)");
                var df = Regex.Matches(text, @"df\s*=\s*k\s*-\s*1\s*-\s*p\s+(\d+)");

                if (calculated.Count >= 2 && critical.Count >= 2 && df.Count >= 2)
                {
                    poisson.ChiSquareStat = double.Parse(calculated[0].Groups[1].Value);
                    poisson.CriticalValue = double.Parse(critical[0].Groups[1].Value);
                    poisson.DegreesOfFreedom = int.Parse(df[0].Groups[1].Value);

                    exp.ChiSquareStat = double.Parse(calculated[1].Groups[1].Value);
                    exp.CriticalValue = double.Parse(critical[1].Groups[1].Value);
                    exp.DegreesOfFreedom = int.Parse(df[1].Groups[1].Value);

                    return (poisson, exp, true);
                }
            }
            catch
            {
                // fall through to fallback
            }
            return (poisson, exp, false);
        }

        // ====================================================================
        // IMPORT #3: SPSS PDF -> parsed Service Time descriptives (cross-check)
        // ====================================================================
        private static (double mean, double sd, bool parsedOk) TryLoadServiceTimeDescriptivesFromSpss(string pdfPath)
        {
            const double fallbackMean = 2.70;
            const double fallbackSd = 0.923;

            string text;
            try
            {
                text = ExtractPdfText(pdfPath);
            }
            catch
            {
                return (fallbackMean, fallbackSd, false);
            }

            try
            {
                // Expected line shape: "Service Time (min) 20 2 5 2.70 .923"
                var m = Regex.Match(text,
                    @"Service Time\s*\(min\)\s+\d+\s+\d+\s+\d+\s+([\d.]+)\s+(\.?\d+\.?\d*)");
                if (m.Success)
                {
                    double mean = double.Parse(m.Groups[1].Value);
                    string sdStr = m.Groups[2].Value;
                    if (sdStr.StartsWith(".")) sdStr = "0" + sdStr;
                    double sd = double.Parse(sdStr);
                    return (mean, sd, true);
                }
            }
            catch
            {
                // fall through to fallback
            }
            return (fallbackMean, fallbackSd, false);
        }

        private static string ExtractPdfText(string pdfPath)
        {
            var sb = new StringBuilder();
            using var document = PdfDocument.Open(pdfPath);
            foreach (var page in document.GetPages())
                sb.AppendLine(page.Text);
            return sb.ToString();
        }

        // ====================================================================
        // Helpers
        // ====================================================================
        private static double TimeAverageCount(List<(double start, double end)> intervals,
                                                 double windowStart, double windowEnd)
        {
            var events = new List<(double time, int delta)>();
            foreach (var (start, end) in intervals)
            {
                events.Add((start, +1));
                events.Add((end, -1));
            }
            events.Sort((a, b) => a.time.CompareTo(b.time));

            double area = 0.0;
            double tPrev = windowStart;
            int n = 0;
            foreach (var (time, delta) in events)
            {
                area += n * (time - tPrev);
                n += delta;
                tPrev = time;
            }
            area += n * (windowEnd - tPrev);

            double totalTime = windowEnd - windowStart;
            return totalTime > 0 ? area / totalTime : 0.0;
        }

        private static double StdDev(List<double> values)
        {
            double mean = values.Average();
            double sumSq = values.Sum(v => (v - mean) * (v - mean));
            return Math.Sqrt(sumSq / (values.Count - 1)); // sample SD, matches SPSS
        }
    }
}