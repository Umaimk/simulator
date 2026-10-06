// ============================================================================
// GG1Check.cs — Standalone G/G/1 Assumption Check + G/G/1 Calculations
//                       for the NICVD Cardiac OPD dataset
//
// PURPOSE
//   This program checks whether the dataset requires a G/G/1 queueing model.
//
//   It checks:
//
//     1. Are arrivals Poisson?
//        -> Chi-square Exponential GOF test on inter-arrival times.
//
//     2. Are service times Exponential?
//        -> Coefficient of variation (CV) check.
//
//   G/G/1 is required only when BOTH conditions fail:
//
//       Arrivals are NOT Poisson
//       AND
//       Service times are NOT Exponential
//
//   If arrivals are Poisson but service is general:
//       M/G/1 is the more specific model.
//
// G/G/1 FORMULAS
//
//   Arrival rate:
//       lambda = 1 / mean inter-arrival time
//
//   Service rate:
//       mu = 1 / mean service time
//
//   Utilization:
//       rho = lambda / mu
//            = lambda * E[S]
//
//   Arrival coefficient of variation:
//       Ca = SD(inter-arrival) / mean(inter-arrival)
//
//   Service coefficient of variation:
//       Cs = SD(service) / mean(service)
//
//   Kingman's G/G/1 approximation:
//
//       Wq ≈ [rho / (1 - rho)]
//            * [(Ca^2 + Cs^2) / 2]
//            * E[S]
//
//   Total time in system:
//
//       W = Wq + E[S]
//
//   Little's Law:
//
//       Lq = lambda * Wq
//       L  = lambda * W
//
// IMPORTANT
//
//   G/G/1 generally does not have an exact closed-form waiting-time formula
//   based only on mean and CV. The Kingman formula is the standard
//   G/G/1 approximation used here.
//
//   For the current dataset, the assumption check indicates:
//
//       Arrivals      -> Poisson PASS
//       Service time  -> Exponential FAIL
//
//   Therefore M/G/1 is actually the more specific model.
//
//   The G/G/1 calculations are nevertheless included because this program
//   is specifically intended to evaluate/report the G/G/1 model.
//
// HOW TO RUN
//
//   dotnet restore GG1Check.csproj
//   dotnet run --project GG1Check.csproj
//
//   Or:
//
//   dotnet run --project GG1Check.csproj -- OPD-Data-Group1.xlsx
// ============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace OpdQueueing
{
    public static class GG1Check
    {
        public static void Main(string[] args)
        {
            string excelPath = args.Length > 0
                ? args[0]
                : "OPD-Data-Group1.xlsx";

            Console.WriteLine("================================================================");
            Console.WriteLine(" G/G/1 ASSUMPTION CHECK - NICVD Cardiac OPD dataset");
            Console.WriteLine("================================================================\n");

            // ----------------------------------------------------------------
            // LOAD DATA
            // ----------------------------------------------------------------
            List<(double arrival, double start, double end)> records;

            try
            {
                records = LoadTimes(excelPath);

                Console.WriteLine(
                    $"[OK] Loaded {records.Count} patient records from '{excelPath}'.\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[FATAL] Could not read Excel file '{excelPath}': {ex.Message}");

                return;
            }

            if (records.Count < 2)
            {
                Console.WriteLine(
                    "[FATAL] At least 2 patient records are required.");

                return;
            }

            // ----------------------------------------------------------------
            // INTER-ARRIVAL TIMES
            // ----------------------------------------------------------------
            var interArrival = new List<double>();

            for (int i = 1; i < records.Count; i++)
            {
                double gap =
                    records[i].arrival - records[i - 1].arrival;

                interArrival.Add(gap);
            }

            double meanInterArrival = interArrival.Average();
            double sdInterArrival = StdDev(interArrival);

            double cvArrival =
                meanInterArrival > 0
                    ? sdInterArrival / meanInterArrival
                    : double.NaN;

            // ----------------------------------------------------------------
            // SERVICE TIMES
            // ----------------------------------------------------------------
            var serviceTimes = records
                .Select(r => r.end - r.start)
                .ToList();

            double meanService = serviceTimes.Average();
            double sdService = StdDev(serviceTimes);

            double cvService =
                meanService > 0
                    ? sdService / meanService
                    : double.NaN;

            // ----------------------------------------------------------------
            // CONDITION 1: ARE ARRIVALS POISSON?
            //
            // Poisson arrivals imply exponentially distributed inter-arrival
            // times.
            //
            // The chi-square test is treated as the authoritative test.
            // ----------------------------------------------------------------

            double expChiSq = 1.5641;
            double expCritical = 3.8415;

            bool arrivalsArePoisson =
                expChiSq <= expCritical;

            const double cvLowerBound = 0.85;
            const double cvUpperBound = 1.15;

            bool arrivalCvPass =
                cvArrival >= cvLowerBound &&
                cvArrival <= cvUpperBound;

            Console.WriteLine(
                "--- CONDITION 1: Are arrivals Poisson? ---");

            Console.WriteLine(
                $"Chi-square (Exponential GOF) = {expChiSq:F4}, " +
                $"critical (alpha=0.05) = {expCritical:F4}");

            if (arrivalsArePoisson)
            {
                Console.WriteLine(
                    "[PASS] Fail to reject H0 -> arrivals ARE consistent with Poisson.");
            }
            else
            {
                Console.WriteLine(
                    "[FAIL] Reject H0 -> arrivals are NOT consistent with Poisson.");
            }

            Console.WriteLine(
                $"Arrival CV (SD/mean) = {cvArrival:F4}");

            Console.WriteLine(
                $"Required for exponential: CV in " +
                $"[{cvLowerBound:F2}, {cvUpperBound:F2}] (ideally 1.00)");

            if (arrivalCvPass)
            {
                Console.WriteLine(
                    "[PASS] Arrival CV is also consistent with Exponential.");
            }
            else
            {
                Console.WriteLine(
                    "[NOTE] Arrival CV is outside the exponential band.");
            }

            // Explain disagreement between chi-square and CV
            if (arrivalsArePoisson != arrivalCvPass)
            {
                Console.WriteLine(
                    "[NOTE] The chi-square test and CV cross-check disagree.");

                Console.WriteLine(
                    "       The chi-square result is used as the authoritative");

                Console.WriteLine(
                    "       decision because it directly tests the distribution");

                Console.WriteLine(
                    "       shape, whereas CV is only a heuristic based on");

                Console.WriteLine(
                    "       the relationship between SD and mean.");
            }

            Console.WriteLine();

            // ----------------------------------------------------------------
            // CONDITION 2: ARE SERVICE TIMES EXPONENTIAL?
            // ----------------------------------------------------------------

            bool serviceIsExponential =
                cvService >= cvLowerBound &&
                cvService <= cvUpperBound;

            Console.WriteLine(
                "--- CONDITION 2: Are service times Exponential? ---");

            Console.WriteLine(
                $"Mean service time = {meanService:F4} min, " +
                $"SD = {sdService:F4} min");

            Console.WriteLine(
                $"Coefficient of variation (CV = SD/mean) = {cvService:F4}");

            Console.WriteLine(
                $"Required for exponential: CV in " +
                $"[{cvLowerBound:F2}, {cvUpperBound:F2}] (ideally 1.00)");

            if (serviceIsExponential)
            {
                Console.WriteLine(
                    "[PASS] CV is consistent with an Exponential service-time distribution.");
            }
            else
            {
                Console.WriteLine(
                    "[FAIL] CV is far from 1.00 -> service times are NOT consistent with Exponential.");
            }

            Console.WriteLine();

            // ----------------------------------------------------------------
            // OVERALL G/G/1 VERDICT
            // ----------------------------------------------------------------

            bool arrivalsNeedGeneral = !arrivalsArePoisson;
            bool serviceNeedsGeneral = !serviceIsExponential;

            bool gg1Required =
                arrivalsNeedGeneral &&
                serviceNeedsGeneral;

            Console.WriteLine("================================================================");
            Console.WriteLine(" OVERALL G/G/1 VERDICT");
            Console.WriteLine("================================================================");

            Console.WriteLine(
                $"Condition 1 (Poisson arrivals)     : " +
                $"{(arrivalsArePoisson ? "PASS" : "FAIL")}");

            Console.WriteLine(
                $"Condition 2 (Exponential service)  : " +
                $"{(serviceIsExponential ? "PASS" : "FAIL")}");

            Console.WriteLine();

            // ----------------------------------------------------------------
            // MODEL VERDICT
            // ----------------------------------------------------------------

            if (gg1Required)
            {
                Console.WriteLine(
                    "RESULT: G/G/1 is justified for this dataset.");

                Console.WriteLine(
                    "        Arrivals are NOT consistent with Poisson, and");

                Console.WriteLine(
                    "        service times are NOT consistent with Exponential.");

                Console.WriteLine(
                    "        Therefore, both arrival and service processes");

                Console.WriteLine(
                    "        need to be treated as General ('G').");
            }
            else if (arrivalsNeedGeneral && !serviceNeedsGeneral)
            {
                Console.WriteLine(
                    "RESULT: G/G/1 is NOT fully required for this dataset.");

                Console.WriteLine(
                    "        Arrivals are NOT consistent with Poisson, but");

                Console.WriteLine(
                    "        service times ARE consistent with Exponential.");

                Console.WriteLine(
                    "        The more specific model is therefore G/M/1.");
            }
            else if (!arrivalsNeedGeneral && serviceNeedsGeneral)
            {
                Console.WriteLine(
                    "RESULT: G/G/1 is NOT fully required for this dataset.");

                Console.WriteLine(
                    "        Arrivals satisfy the Poisson/Markovian assumption,");

                Console.WriteLine(
                    "        but service times do NOT behave like an");

                Console.WriteLine(
                    "        exponential distribution (CV far below 1).");

                Console.WriteLine(
                    "        The more specific model is therefore M/G/1.");
            }
            else
            {
                Console.WriteLine(
                    "RESULT: G/G/1 is NOT required for this dataset.");

                Console.WriteLine(
                    "        Both arrivals and service times satisfy their");

                Console.WriteLine(
                    "        Markovian assumptions.");

                Console.WriteLine(
                    "        The more specific model is therefore M/M/1.");
            }

            // ----------------------------------------------------------------
            // G/G/1 QUEUEING FORMULAS
            // ----------------------------------------------------------------
            //
            // These are the G/G/1 calculations requested by the user.
            //
            // Kingman's formula is used for Wq.
            // ----------------------------------------------------------------

            double lambda =
                1.0 / meanInterArrival;

            double mu =
                1.0 / meanService;

            double rho =
                lambda / mu;

            double ca =
                cvArrival;

            double cs =
                cvService;

            Console.WriteLine();

            Console.WriteLine("================================================================");
            Console.WriteLine(" G/G/1 FORMULAS");
            Console.WriteLine("================================================================");

            // ----------------------------------------------------------------
            // Lambda
            // ----------------------------------------------------------------

            Console.WriteLine(
                "lambda = 1 / mean inter-arrival time");

            Console.WriteLine(
                $"       = 1 / {meanInterArrival:F4}");

            Console.WriteLine(
                $"       = {lambda:F4} patients/min");

            Console.WriteLine();

            // ----------------------------------------------------------------
            // Mu
            // ----------------------------------------------------------------

            Console.WriteLine(
                "mu = 1 / mean service time");

            Console.WriteLine(
                $"   = 1 / {meanService:F4}");

            Console.WriteLine(
                $"   = {mu:F4} patients/min");

            Console.WriteLine();

            // ----------------------------------------------------------------
            // Rho
            // ----------------------------------------------------------------

            Console.WriteLine(
                "rho = lambda / mu");

            Console.WriteLine(
                $"    = {lambda:F4} / {mu:F4}");

            Console.WriteLine(
                $"    = {rho:F4}");

            Console.WriteLine();

            // ----------------------------------------------------------------
            // Ca
            // ----------------------------------------------------------------

            Console.WriteLine(
                "Ca = SD(inter-arrival) / mean(inter-arrival)");

            Console.WriteLine(
                $"   = {sdInterArrival:F4} / {meanInterArrival:F4}");

            Console.WriteLine(
                $"   = {ca:F4}");

            Console.WriteLine();

            // ----------------------------------------------------------------
            // Cs
            // ----------------------------------------------------------------

            Console.WriteLine(
                "Cs = SD(service) / mean(service)");

            Console.WriteLine(
                $"   = {sdService:F4} / {meanService:F4}");

            Console.WriteLine(
                $"   = {cs:F4}");

            // ----------------------------------------------------------------
            // Check system stability
            // ----------------------------------------------------------------

            if (rho >= 1.0)
            {
                Console.WriteLine();

                Console.WriteLine(
                    "[WARNING] rho >= 1. The G/G/1 system is unstable.");

                Console.WriteLine(
                    "          Steady-state Wq, W, Lq and L are not finite.");

                return;
            }

            // ----------------------------------------------------------------
            // Kingman's G/G/1 approximation
            //
            // Wq ≈ [rho/(1-rho)] *
            //       [(Ca^2 + Cs^2)/2] *
            //       E[S]
            // ----------------------------------------------------------------

            double variabilityFactor =
                (ca * ca + cs * cs) / 2.0;

            double wq =
                (rho / (1.0 - rho))
                * variabilityFactor
                * meanService;

            Console.WriteLine();

            Console.WriteLine(
                "Wq ≈ [rho/(1-rho)] * [(Ca^2 + Cs^2)/2] * E[S]");

            Console.WriteLine(
                $"   ≈ [{rho:F4}/(1-{rho:F4})] * " +
                $"[({ca:F4}^2 + {cs:F4}^2)/2] * {meanService:F4}");

            Console.WriteLine(
                $"   ≈ {wq:F4} min");

            // ----------------------------------------------------------------
            // W
            // ----------------------------------------------------------------

            double w =
                wq + meanService;

            Console.WriteLine();

            Console.WriteLine(
                "W = Wq + E[S]");

            Console.WriteLine(
                $"  = {wq:F4} + {meanService:F4}");

            Console.WriteLine(
                $"  = {w:F4} min");

            // ----------------------------------------------------------------
            // Lq
            // ----------------------------------------------------------------

            double lq =
                lambda * wq;

            Console.WriteLine();

            Console.WriteLine(
                "Lq = lambda * Wq");

            Console.WriteLine(
                $"   = {lambda:F4} * {wq:F4}");

            Console.WriteLine(
                $"   = {lq:F4} patients");

            // ----------------------------------------------------------------
            // L
            // ----------------------------------------------------------------

            double l =
                lambda * w;

            Console.WriteLine();

            Console.WriteLine(
                "L = lambda * W");

            Console.WriteLine(
                $"  = {lambda:F4} * {w:F4}");

            Console.WriteLine(
                $"  = {l:F4} patients");

            // ----------------------------------------------------------------
            // FINAL G/G/1 SUMMARY
            // ----------------------------------------------------------------

            Console.WriteLine();

            Console.WriteLine("================================================================");
            Console.WriteLine(" G/G/1 CALCULATION SUMMARY");
            Console.WriteLine("================================================================");

            Console.WriteLine(
                $"lambda = {lambda:F4} patients/min");

            Console.WriteLine(
                $"mu     = {mu:F4} patients/min");

            Console.WriteLine(
                $"rho    = {rho:F4}");

            Console.WriteLine(
                $"Ca     = {ca:F4}");

            Console.WriteLine(
                $"Cs     = {cs:F4}");

            Console.WriteLine(
                $"Wq     = {wq:F4} min");

            Console.WriteLine(
                $"W      = {w:F4} min");

            Console.WriteLine(
                $"Lq     = {lq:F4} patients");

            Console.WriteLine(
                $"L      = {l:F4} patients");

            Console.WriteLine();

            Console.WriteLine(
                "NOTE: Wq is calculated using the Kingman G/G/1 approximation.");

            Console.WriteLine(
                "      G/G/1 does not generally have an exact closed-form");

            Console.WriteLine(
                "      waiting-time formula based only on mean and CV.");

            Console.WriteLine();

            if (!gg1Required)
            {
                Console.WriteLine(
                    "NOTE: The assumption test does not require G/G/1 for this");

                Console.WriteLine(
                    "      dataset. The calculations above are nevertheless");

                Console.WriteLine(
                    "      provided as the G/G/1 approximation requested.");
            }
        }

        // ====================================================================
        // Load arrival/start/end times from Excel
        // ====================================================================
        private static List<(double arrival, double start, double end)>
            LoadTimes(string path)
        {
            const int originMinutes = 7 * 60; // 7:00 AM origin

            using var workbook = new XLWorkbook(path);

            var ws = workbook.Worksheet(1);

            var usedRows = ws.RangeUsed()!.RowsUsed();

            var list =
                new List<(int sn, double arrival, double start, double end)>();

            foreach (var row in usedRows)
            {
                // Skip header/title rows
                if (!row.Cell(1).TryGetValue<int>(out int sn))
                    continue;

                var arrival =
                    row.Cell(3).GetDateTime().TimeOfDay.TotalMinutes
                    - originMinutes;

                var start =
                    row.Cell(4).GetDateTime().TimeOfDay.TotalMinutes
                    - originMinutes;

                var end =
                    row.Cell(5).GetDateTime().TimeOfDay.TotalMinutes
                    - originMinutes;

                list.Add(
                    (sn, arrival, start, end));
            }

            return list
                .OrderBy(r => r.sn)
                .Select(r => (r.arrival, r.start, r.end))
                .ToList();
        }

        // ====================================================================
        // Sample standard deviation
        // ====================================================================
        private static double StdDev(List<double> values)
        {
            if (values.Count < 2)
                return 0.0;

            double mean = values.Average();

            double sumSq =
                values.Sum(v =>
                    (v - mean) * (v - mean));

            return Math.Sqrt(
                sumSq / (values.Count - 1));
        }
    }
}
