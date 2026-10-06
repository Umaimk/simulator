// ============================================================================
// Mm1Check.cs — Standalone M/M/1 Assumption Check for the OPD dataset
//
// PURPOSE
//   This file is separate from Program.cs on purpose. Program.cs already
//   computes M/M/1 as a "reference only" comparison alongside the selected
//   M/G/1 model. This file instead answers one narrow question by itself:
//
//       "Does this dataset actually qualify as M/M/1?"
//
//   It re-checks the two conditions M/M/1 requires:
//     1. Arrivals are Poisson (Markovian)        -> from the chi-square result
//     2. Service times are Exponential (Markovian) -> via the coefficient of
//        variation (CV) test: CV = SD/mean must be close to 1.00 for a true
//        exponential distribution. No separate chi-square test for service
//        time was supplied in your files, so CV is the best available check.
//
//   It prints a clear PASS/FAIL verdict for each condition, an overall
//   verdict for whether M/M/1 is justified, and then computes the M/M/1
//   formulas anyway (clearly labeled), so you can see the numbers either way.
//
// HOW TO RUN THIS FILE ON ITS OWN
//   This project (OpdQueueingAnalysis.csproj) currently compiles every .cs
//   file in the folder into ONE executable, which means Program.cs (with its
//   own Main) and this file would collide if both define a Main method in
//   the same build. To run ONLY this check:
//
//   Option A - Temporary rename (simplest):
//     1. Rename Program.cs to Program.cs.bak (so it's excluded from the build)
//     2. Run: dotnet run
//     3. Rename Program.cs.bak back to Program.cs when done
//
//   Option B - Separate project folder (cleaner, recommended if you'll run
//   this often): create a new folder (e.g. Mm1CheckApp), copy this file and
//   a copy of OpdQueueingAnalysis.csproj into it, then run `dotnet run` from
//   inside that new folder. Keep OPD-Data-Group1.xlsx and OPD-SPSS.pdf
//   next to it too (same filenames as before).
//
// This file imports the SAME two source files as Program.cs:
//   - OPD-Data-Group1.xlsx  (ClosedXML)  -> service times
//   - OPD-SPSS.pdf          (PdfPig)     -> cross-check for service time mean/SD
// The chi-square Poisson/Exponential-arrivals result is taken from the
// transcribed PDF values (same figures used in Program.cs), since that test
// was already performed on arrivals, not service time.
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
    public static class Mm1Check
    {
        public static void Main(string[] args)
        {
            string excelPath = args.Length > 0 ? args[0] : "OPD-Data-Group1.xlsx";
            string spssPdfPath = args.Length > 1 ? args[1] : "OPD-SPSS.pdf";

            Console.WriteLine("================================================================");
            Console.WriteLine(" M/M/1 ASSUMPTION CHECK - NICVD Cardiac OPD dataset");
            Console.WriteLine("================================================================\n");

            // ----------------------------------------------------------------
            // Load the raw data (same loader logic as Program.cs)
            // ----------------------------------------------------------------
            List<(double arrival, double start, double end)> records;
            try
            {
                records = LoadTimes(excelPath);
                Console.WriteLine($"[OK] Loaded {records.Count} patient records from '{excelPath}'.\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FATAL] Could not read Excel file '{excelPath}': {ex.Message}");
                return;
            }

            var serviceTimes = records.Select(r => r.end - r.start).ToList();
            double meanService = serviceTimes.Average();
            double sdService = StdDev(serviceTimes);
            double cv = sdService / meanService;

            var interArrival = new List<double>();
            for (int i = 1; i < records.Count; i++)
                interArrival.Add(records[i].arrival - records[i - 1].arrival);
            double meanInterArrival = interArrival.Average();

            // ----------------------------------------------------------------
            // Cross-check service time stats against the SPSS PDF
            // ----------------------------------------------------------------
            var (spssMean, spssSd, spssParsed) = TryLoadServiceTimeFromSpss(spssPdfPath);
            Console.WriteLine("--- Service time (cross-checked against SPSS PDF) ---");
            Console.WriteLine($"Mean service time = {meanService:F4} min, SD = {sdService:F4} min");
            if (spssParsed)
                Console.WriteLine($"SPSS reports      = mean {spssMean:F2}, SD {spssSd:F3}  " +
                                   (Math.Abs(spssMean - meanService) < 0.05 && Math.Abs(spssSd - sdService) < 0.05
                                       ? "(matches)" : "(MISMATCH - check data)"));
            else
                Console.WriteLine("[WARNING] Could not parse SPSS PDF; using Excel-derived values only.");
            Console.WriteLine();

            // ----------------------------------------------------------------
            // CONDITION 1: Arrivals Poisson? (from the externally-run chi-square test)
            // ----------------------------------------------------------------
            // Transcribed directly from Chi_Square_Analysison_arrival_time.pdf
            double poissonChiSq = 5.0859, poissonCritical = 12.5916;
            bool poissonPass = poissonChiSq <= poissonCritical;

            Console.WriteLine("--- CONDITION 1: Are arrivals Poisson? ---");
            Console.WriteLine($"Chi-square = {poissonChiSq}, critical (alpha=0.05) = {poissonCritical}");
            Console.WriteLine(poissonPass
                ? "[PASS] Fail to reject H0 -> arrivals ARE consistent with Poisson.\n"
                : "[FAIL] Reject H0 -> arrivals are NOT consistent with Poisson.\n");

            // ----------------------------------------------------------------
            // CONDITION 2: Service times Exponential? (via coefficient of variation)
            // ----------------------------------------------------------------
            // An exponential distribution must have CV = SD/mean = 1.00 exactly.
            // No formal chi-square test for service time was supplied, so this
            // program uses a standard tolerance band around 1.00 to judge it.
            const double cvLowerBound = 0.85;
            const double cvUpperBound = 1.15;
            bool servicePass = cv >= cvLowerBound && cv <= cvUpperBound;

            Console.WriteLine("--- CONDITION 2: Are service times Exponential? ---");
            Console.WriteLine($"Coefficient of variation (CV = SD/mean) = {cv:F4}");
            Console.WriteLine($"Required for exponential: CV in [{cvLowerBound:F2}, {cvUpperBound:F2}] (ideally 1.00)");
            Console.WriteLine(servicePass
                ? "[PASS] CV is close enough to 1.00 -> service times ARE consistent with Exponential.\n"
                : "[FAIL] CV is far from 1.00 -> service times are NOT consistent with Exponential.\n");

            // ----------------------------------------------------------------
            // OVERALL VERDICT
            // ----------------------------------------------------------------
            bool mm1Justified = poissonPass && servicePass;
            Console.WriteLine("================================================================");
            Console.WriteLine(" OVERALL M/M/1 VERDICT");
            Console.WriteLine("================================================================");
            Console.WriteLine($"Condition 1 (Poisson arrivals)     : {(poissonPass ? "PASS" : "FAIL")}");
            Console.WriteLine($"Condition 2 (Exponential service)  : {(servicePass ? "PASS" : "FAIL")}");
            Console.WriteLine();
            if (mm1Justified)
            {
                Console.WriteLine("RESULT: M/M/1 is statistically justified for this dataset.");
            }
            else
            {
                Console.WriteLine("RESULT: M/M/1 is NOT fully justified for this dataset.");
                Console.WriteLine("        Arrivals satisfy the Markovian assumption, but service time");
                Console.WriteLine("        does not behave like an exponential distribution (CV far below 1).");
                Console.WriteLine("        The better-supported model is M/G/1 (see Program.cs / main report).");
            }
            Console.WriteLine();

            // ----------------------------------------------------------------
            // Compute M/M/1 formulas regardless, for reference
            // ----------------------------------------------------------------
            double lambda = 1.0 / meanInterArrival;
            double mu = 1.0 / meanService;
            double rho = lambda / mu;

            Console.WriteLine("--- M/M/1 formulas, computed regardless of the verdict above ---");
            Console.WriteLine($"lambda = {lambda:F4} patients/min");
            Console.WriteLine($"mu     = {mu:F4} patients/min");
            Console.WriteLine($"rho    = {rho:F4}");

            if (rho < 1)
            {
                double L = rho / (1 - rho);
                double Lq = (rho * rho) / (1 - rho);
                double W = 1.0 / (mu - lambda);
                double Wq = rho / (mu - lambda);

                Console.WriteLine($"L  = rho/(1-rho)     = {L:F4} patients");
                Console.WriteLine($"Lq = rho^2/(1-rho)   = {Lq:F4} patients");
                Console.WriteLine($"W  = 1/(mu-lambda)   = {W:F4} min");
                Console.WriteLine($"Wq = rho/(mu-lambda) = {Wq:F4} min");
            }
            else
            {
                Console.WriteLine("System is unstable (rho >= 1) -> M/M/1 formulas do not apply.");
            }

            Console.WriteLine();
            Console.WriteLine(mm1Justified
                ? "NOTE: Even though M/M/1 passed both checks here, always confirm against"
                : "NOTE: Because M/M/1 failed the service-time check, these M/M/1 numbers");
            Console.WriteLine(mm1Justified
                ? "      the full chi-square test results in your report before relying on it."
                : "      should be treated as a reference/comparison only, not the real model.");
        }

        // ====================================================================
        // Load arrival/start/end times from the Excel dataset
        // ====================================================================
        private static List<(double arrival, double start, double end)> LoadTimes(string path)
        {
            const int originMinutes = 7 * 60; // 7:00 AM origin, matches SPSS numbering

            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheet(1);
            var usedRows = ws.RangeUsed()!.RowsUsed();

            var list = new List<(int sn, double arrival, double start, double end)>();
            foreach (var row in usedRows)
            {
                if (!row.Cell(1).TryGetValue<int>(out int sn)) continue; // skip header/title rows

                var arrival = row.Cell(3).GetDateTime().TimeOfDay.TotalMinutes - originMinutes;
                var start = row.Cell(4).GetDateTime().TimeOfDay.TotalMinutes - originMinutes;
                var end = row.Cell(5).GetDateTime().TimeOfDay.TotalMinutes - originMinutes;
                list.Add((sn, arrival, start, end));
            }

            return list.OrderBy(r => r.sn)
                       .Select(r => (r.arrival, r.start, r.end))
                       .ToList();
        }

        // ====================================================================
        // Cross-check against the SPSS PDF (same approach as Program.cs)
        // ====================================================================
        private static (double mean, double sd, bool parsedOk) TryLoadServiceTimeFromSpss(string pdfPath)
        {
            try
            {
                var sb = new StringBuilder();
                using (var document = PdfDocument.Open(pdfPath))
                    foreach (var page in document.GetPages())
                        sb.AppendLine(page.Text);

                var m = Regex.Match(sb.ToString(),
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
                // fall through
            }
            return (0, 0, false);
        }

        private static double StdDev(List<double> values)
        {
            double mean = values.Average();
            double sumSq = values.Sum(v => (v - mean) * (v - mean));
            return Math.Sqrt(sumSq / (values.Count - 1));
        }
    }
}
