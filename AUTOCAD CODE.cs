// VisJsonNet.cs  (C# 7.3 compatible, AutoCAD 2022, x64)
// References: acmgd.dll, acdbmgd.dll, accoremgd.dll  (Copy Local = False)
//
// ✅ Generates TWO JSON files:
// 1) vis_export_visibility.json  -> ONLY dynamic blocks that have Visibility
//    Grouped by: (baseName + visibility)
//    Includes: tableCount, chairCount, length_ft, width_ft (median), distance1..4_ft (median)
// 2) vis_export_all.json         -> ALL modelspace block references
//    Grouped by: (baseName)
//    Includes: totalCount, dynamicCount, length_ft, width_ft (median), distance1..4_ft (median)
//
// ✅ Fixes *U### names via DynamicBlockTableRecord for baseName.
// ✅ Size in FEET (rotation-safe):
//    - Computes extents in evaluated block definition (br.BlockTableRecord) local coords,
//      then applies br.ScaleFactors and INSUNITS → feet conversion.
//
// Output folder:
//   C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS\

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

namespace DwgExtractPlugin
{
    public class VisJsonNet : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("PINGNET")]
        public void PingNet()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc != null) doc.Editor.WriteMessage("\n✅ PINGNET OK\n");
        }

        [CommandMethod("VISJSONNET")]
        public void VisJsonNetCommand()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
                throw new System.Exception("No active document.");

            Database db = doc.Database;

            string dwgFullPath = db.Filename;
            if (string.IsNullOrWhiteSpace(dwgFullPath))
                dwgFullPath = doc.Name;

            string dwgFileName = Path.GetFileName(dwgFullPath);

            string outDir = @"C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS";
            Directory.CreateDirectory(outDir);

            string outJsonVisibility = Path.Combine(outDir, "vis_export_visibility.json");
            string outJsonAll = Path.Combine(outDir, "vis_export_all.json");

            double unitToFt = GetScaleToFeet(db);

            var visGrouped = new Dictionary<string, VisAgg>(StringComparer.OrdinalIgnoreCase);
            var allGrouped = new Dictionary<string, AllAgg>(StringComparer.OrdinalIgnoreCase);

            int totalInserts = 0;
            int dynamicProcessed = 0;
            int visKept = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                        continue;

                    totalInserts++;

                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

                    string baseName = Safe(GetBaseName(tr, br));
                    if (string.IsNullOrWhiteSpace(baseName))
                        baseName = Safe(br.Name);

                    // Rotation-safe size in feet
                    double lenFt, widFt;
                    GetBlockSizeFeet(tr, br, unitToFt, out lenFt, out widFt);

                    // Distance params (in feet) — only dynamic blocks will usually have these
                    double d1Ft, d2Ft, d3Ft, d4Ft;
                    GetDistanceParamsFeet(br, unitToFt, out d1Ft, out d2Ft, out d3Ft, out d4Ft);

                    // ---------- (2) ALL JSON aggregation ----------
                    {
                        AllAgg a;
                        if (!allGrouped.TryGetValue(baseName, out a))
                        {
                            a = new AllAgg
                            {
                                BlockName = baseName,
                                TotalCount = 0,
                                DynamicCount = 0
                            };
                            allGrouped[baseName] = a;
                        }

                        a.TotalCount += 1;
                        if (br.IsDynamicBlock) a.DynamicCount += 1;

                        if (lenFt > 0) a.LenFts.Add(lenFt);
                        if (widFt > 0) a.WidFts.Add(widFt);

                        if (d1Ft > 0) a.D1Fts.Add(d1Ft);
                        if (d2Ft > 0) a.D2Fts.Add(d2Ft);
                        if (d3Ft > 0) a.D3Fts.Add(d3Ft);
                        if (d4Ft > 0) a.D4Fts.Add(d4Ft);
                    }

                    // ---------- (1) VISIBILITY JSON aggregation ----------
                    string vis = "";
                    if (br.IsDynamicBlock)
                    {
                        dynamicProcessed++;
                        vis = GetVisibilityState(br);
                    }

                    if (!string.IsNullOrWhiteSpace(vis))
                    {
                        visKept++;

                        int chairCount = GetChairCountFromDynProps(br);
                        if (chairCount <= 0)
                            chairCount = CountNestedChairsFromEvaluatedBTR(tr, br);

                        string key = baseName + "||" + vis;

                        VisAgg v;
                        if (!visGrouped.TryGetValue(key, out v))
                        {
                            v = new VisAgg
                            {
                                TableName = baseName,
                                Visibility = vis,
                                TableCount = 0,
                                ChairCount = 0
                            };
                            visGrouped[key] = v;
                        }

                        v.TableCount += 1;
                        v.ChairCount += chairCount;

                        if (lenFt > 0) v.LenFts.Add(lenFt);
                        if (widFt > 0) v.WidFts.Add(widFt);

                        if (d1Ft > 0) v.D1Fts.Add(d1Ft);
                        if (d2Ft > 0) v.D2Fts.Add(d2Ft);
                        if (d3Ft > 0) v.D3Fts.Add(d3Ft);
                        if (d4Ft > 0) v.D4Fts.Add(d4Ft);
                    }
                }

                tr.Commit();
            }

            // finalize medians
            foreach (var v in visGrouped.Values)
            {
                v.LengthFt = Median(v.LenFts);
                v.WidthFt = Median(v.WidFts);

                v.Distance1Ft = Median(v.D1Fts);
                v.Distance2Ft = Median(v.D2Fts);
                v.Distance3Ft = Median(v.D3Fts);
                v.Distance4Ft = Median(v.D4Fts);
            }

            foreach (var a in allGrouped.Values)
            {
                a.LengthFt = Median(a.LenFts);
                a.WidthFt = Median(a.WidFts);

                a.Distance1Ft = Median(a.D1Fts);
                a.Distance2Ft = Median(a.D2Fts);
                a.Distance3Ft = Median(a.D3Fts);
                a.Distance4Ft = Median(a.D4Fts);
            }

            // Write VISIBILITY JSON
            string jsonVis = BuildVisibilityJson(dwgFileName, totalInserts, dynamicProcessed, visKept, visGrouped.Values.ToList());
            File.WriteAllText(outJsonVisibility, jsonVis, Encoding.UTF8);

            // Write ALL JSON
            string jsonAll = BuildAllJson(dwgFileName, totalInserts, allGrouped.Values.ToList());
            File.WriteAllText(outJsonAll, jsonAll, Encoding.UTF8);

            var ddoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (ddoc != null)
            {
                ddoc.Editor.WriteMessage("\n✅ VISJSONNET wrote:\n");
                ddoc.Editor.WriteMessage("  - " + outJsonVisibility + "\n");
                ddoc.Editor.WriteMessage("  - " + outJsonAll + "\n");
            }
        }

        // ---------------- models ----------------

        private class VisAgg
        {
            public string TableName;
            public string Visibility;
            public int TableCount;
            public int ChairCount;

            public readonly List<double> LenFts = new List<double>();
            public readonly List<double> WidFts = new List<double>();

            public readonly List<double> D1Fts = new List<double>();
            public readonly List<double> D2Fts = new List<double>();
            public readonly List<double> D3Fts = new List<double>();
            public readonly List<double> D4Fts = new List<double>();

            public double LengthFt;
            public double WidthFt;

            public double Distance1Ft;
            public double Distance2Ft;
            public double Distance3Ft;
            public double Distance4Ft;
        }

        private class AllAgg
        {
            public string BlockName;
            public int TotalCount;
            public int DynamicCount;

            public readonly List<double> LenFts = new List<double>();
            public readonly List<double> WidFts = new List<double>();

            public readonly List<double> D1Fts = new List<double>();
            public readonly List<double> D2Fts = new List<double>();
            public readonly List<double> D3Fts = new List<double>();
            public readonly List<double> D4Fts = new List<double>();

            public double LengthFt;
            public double WidthFt;

            public double Distance1Ft;
            public double Distance2Ft;
            public double Distance3Ft;
            public double Distance4Ft;
        }

        // ---------------- JSON builders ----------------

        private static string BuildVisibilityJson(string dwg, int total, int dynProcessed, int kept, List<VisAgg> items)
        {
            items = items
                .OrderBy(x => x.TableName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.Visibility, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"dwg\": ").Append(JsonStr(dwg)).Append(",\n");
            sb.Append("  \"totalInserts\": ").Append(total.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            sb.Append("  \"dynamicProcessed\": ").Append(dynProcessed.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            sb.Append("  \"visibilityKept\": ").Append(kept.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            sb.Append("  \"items\": [\n");

            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                sb.Append("    {\n");
                sb.Append("      \"tableName\": ").Append(JsonStr(it.TableName)).Append(",\n");
                sb.Append("      \"visibility\": ").Append(JsonStr(it.Visibility)).Append(",\n");
                sb.Append("      \"tableCount\": ").Append(it.TableCount.ToString(CultureInfo.InvariantCulture)).Append(",\n");
                sb.Append("      \"chairCount\": ").Append(it.ChairCount.ToString(CultureInfo.InvariantCulture)).Append(",\n");
                sb.Append("      \"length_ft\": ").Append(Num(it.LengthFt)).Append(",\n");
                sb.Append("      \"width_ft\": ").Append(Num(it.WidthFt)).Append(",\n");
                sb.Append("      \"distance1_ft\": ").Append(Num(it.Distance1Ft)).Append(",\n");
                sb.Append("      \"distance2_ft\": ").Append(Num(it.Distance2Ft)).Append(",\n");
                sb.Append("      \"distance3_ft\": ").Append(Num(it.Distance3Ft)).Append(",\n");
                sb.Append("      \"distance4_ft\": ").Append(Num(it.Distance4Ft)).Append("\n");
                sb.Append("    }");
                if (i < items.Count - 1) sb.Append(",");
                sb.Append("\n");
            }

            sb.Append("  ]\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        private static string BuildAllJson(string dwg, int total, List<AllAgg> items)
        {
            items = items
                .OrderBy(x => x.BlockName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"dwg\": ").Append(JsonStr(dwg)).Append(",\n");
            sb.Append("  \"totalInserts\": ").Append(total.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            sb.Append("  \"items\": [\n");

            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                sb.Append("    {\n");
                sb.Append("      \"blockName\": ").Append(JsonStr(it.BlockName)).Append(",\n");
                sb.Append("      \"totalCount\": ").Append(it.TotalCount.ToString(CultureInfo.InvariantCulture)).Append(",\n");
                sb.Append("      \"dynamicCount\": ").Append(it.DynamicCount.ToString(CultureInfo.InvariantCulture)).Append(",\n");
                sb.Append("      \"length_ft\": ").Append(Num(it.LengthFt)).Append(",\n");
                sb.Append("      \"width_ft\": ").Append(Num(it.WidthFt)).Append(",\n");
                sb.Append("      \"distance1_ft\": ").Append(Num(it.Distance1Ft)).Append(",\n");
                sb.Append("      \"distance2_ft\": ").Append(Num(it.Distance2Ft)).Append(",\n");
                sb.Append("      \"distance3_ft\": ").Append(Num(it.Distance3Ft)).Append(",\n");
                sb.Append("      \"distance4_ft\": ").Append(Num(it.Distance4Ft)).Append("\n");
                sb.Append("    }");
                if (i < items.Count - 1) sb.Append(",");
                sb.Append("\n");
            }

            sb.Append("  ]\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        // ---------------- common helpers ----------------

        private static string JsonStr(string s)
        {
            if (s == null) s = "";
            s = s.Replace("\\", "\\\\").Replace("\"", "\\\"");
            s = s.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
            return "\"" + s + "\"";
        }

        private static string Safe(string s) { return s ?? ""; }

        private static string Num(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v) || v <= 0) return "0";
            return v.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static double Median(List<double> xs)
        {
            if (xs == null || xs.Count == 0) return 0;
            xs.Sort();
            int n = xs.Count;
            if (n % 2 == 1) return xs[n / 2];
            return 0.5 * (xs[n / 2 - 1] + xs[n / 2]);
        }

        private static double GetScaleToFeet(Database db)
        {
            try
            {
                UnitsValue u = db.Insunits;
                switch (u)
                {
                    case UnitsValue.Feet: return 1.0;
                    case UnitsValue.Inches: return 1.0 / 12.0;
                    case UnitsValue.Millimeters: return 0.003280839895013123;
                    case UnitsValue.Centimeters: return 0.03280839895013123;
                    case UnitsValue.Meters: return 3.280839895013123;
                    case UnitsValue.Kilometers: return 3280.839895013123;
                    case UnitsValue.Yards: return 3.0;
                    case UnitsValue.Miles: return 5280.0;
                    case UnitsValue.Undefined:
                        // If your unitless DWGs are actually mm, use: 0.003280839895013123
                        return 1.0;
                }
                return 1.0;
            }
            catch { return 1.0; }
        }

        // ---------------- block naming / visibility / chairs ----------------

        private static string GetBaseName(Transaction tr, BlockReference br)
        {
            try
            {
                if (br.IsDynamicBlock)
                {
                    ObjectId baseBtrId = br.DynamicBlockTableRecord;
                    if (!baseBtrId.IsNull)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(baseBtrId, OpenMode.ForRead);
                        if (btr != null && !string.IsNullOrEmpty(btr.Name))
                            return btr.Name;
                    }
                }
            }
            catch { }
            return br.Name;
        }

        private static string GetVisibilityState(BlockReference br)
        {
            try
            {
                if (!br.IsDynamicBlock) return "";
                var props = br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty p in props)
                {
                    if (p == null) continue;
                    var n = (p.PropertyName ?? "").ToUpperInvariant();
                    if (n.Contains("VISIBILITY"))
                        return (p.Value != null) ? p.Value.ToString() : "";
                }
            }
            catch { }
            return "";
        }

        private static int GetChairCountFromDynProps(BlockReference br)
        {
            try
            {
                if (!br.IsDynamicBlock) return 0;

                var props = br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty p in props)
                {
                    if (p == null) continue;

                    string n = (p.PropertyName ?? "").ToUpperInvariant();
                    if (!(n.Contains("CHAIR") || n.Contains("SEAT") || n.Contains("QTY") || n.Contains("COUNT")))
                        continue;

                    string val = (p.Value != null) ? p.Value.ToString() : "";

                    int parsed;
                    if (int.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                        return parsed;

                    double d;
                    if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                        return (int)Math.Round(d);
                }
            }
            catch { }
            return 0;
        }

        private static bool LooksLikeChair(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return name.IndexOf("CHAIR", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("SEAT", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("CHR", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int CountNestedChairsFromEvaluatedBTR(Transaction tr, BlockReference br)
        {
            try
            {
                ObjectId evalBtrId = br.BlockTableRecord;
                if (evalBtrId.IsNull) return 0;

                var evalBtr = (BlockTableRecord)tr.GetObject(evalBtrId, OpenMode.ForRead);
                if (evalBtr == null) return 0;

                int chairCount = 0;

                foreach (ObjectId cid in evalBtr)
                {
                    if (!cid.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                        continue;

                    var child = (BlockReference)tr.GetObject(cid, OpenMode.ForRead);

                    string childBase = Safe(GetBaseName(tr, child));
                    string childRaw = Safe(child.Name);

                    if (LooksLikeChair(childBase) || LooksLikeChair(childRaw))
                        chairCount++;
                }

                return chairCount;
            }
            catch { return 0; }
        }

        // ---------------- Distance1..Distance4 extraction ----------------

        private static void GetDistanceParamsFeet(BlockReference br, double unitToFt,
            out double d1Ft, out double d2Ft, out double d3Ft, out double d4Ft)
        {
            d1Ft = d2Ft = d3Ft = d4Ft = 0;

            try
            {
                if (!br.IsDynamicBlock) return;

                var map = new Dictionary<int, double>();

                var props = br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty p in props)
                {
                    if (p == null) continue;
                    string name = (p.PropertyName ?? "").Trim();
                    if (name.Length == 0) continue;

                    string up = name.ToUpperInvariant();
                    if (!up.Contains("DISTANCE")) continue;

                    int idx = ParseTrailingIndex(up); // Distance1 -> 1
                    if (idx <= 0) continue;

                    double val;
                    if (TryGetDouble(p.Value, out val))
                    {
                        val = Math.Abs(val) * unitToFt;
                        map[idx] = val;
                    }
                }

                if (map.ContainsKey(1)) d1Ft = map[1];
                if (map.ContainsKey(2)) d2Ft = map[2];
                if (map.ContainsKey(3)) d3Ft = map[3];
                if (map.ContainsKey(4)) d4Ft = map[4];
            }
            catch
            {
                d1Ft = d2Ft = d3Ft = d4Ft = 0;
            }
        }

        private static int ParseTrailingIndex(string sUpper)
        {
            int n = 0;
            int mul = 1;
            bool any = false;

            for (int i = sUpper.Length - 1; i >= 0; i--)
            {
                char c = sUpper[i];
                if (c >= '0' && c <= '9')
                {
                    any = true;
                    n += (c - '0') * mul;
                    mul *= 10;
                }
                else break;
            }

            return any ? n : 0;
        }

        private static bool TryGetDouble(object v, out double d)
        {
            d = 0;
            if (v == null) return false;

            try
            {
                if (v is double) { d = (double)v; return true; }
                if (v is float) { d = (float)v; return true; }
                if (v is int) { d = (int)v; return true; }
                if (v is long) { d = (long)v; return true; }
                if (v is short) { d = (short)v; return true; }

                string s = v.ToString();
                if (string.IsNullOrWhiteSpace(s)) return false;

                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)
                    || double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d);
            }
            catch
            {
                return false;
            }
        }

        // ---------------- size in feet (rotation-safe) ----------------

        private static void GetBlockSizeFeet(Transaction tr, BlockReference br, double unitToFt, out double lengthFt, out double widthFt)
        {
            lengthFt = 0;
            widthFt = 0;

            try
            {
                ObjectId evalBtrId = br.BlockTableRecord;
                if (evalBtrId.IsNull) return;

                var evalBtr = (BlockTableRecord)tr.GetObject(evalBtrId, OpenMode.ForRead);
                if (evalBtr == null) return;

                bool has = false;
                Extents3d ext = new Extents3d();

                foreach (ObjectId eid in evalBtr)
                {
                    Entity ent = tr.GetObject(eid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    try
                    {
                        Extents3d eext = ent.GeometricExtents;
                        if (!has) { ext = eext; has = true; }
                        else ext.AddExtents(eext);
                    }
                    catch { }
                }

                if (!has) return;

                double dx = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
                double dy = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);

                double sx = 1.0, sy = 1.0;
                try
                {
                    sx = Math.Abs(br.ScaleFactors.X);
                    sy = Math.Abs(br.ScaleFactors.Y);
                }
                catch { }

                dx *= sx;
                dy *= sy;

                dx *= unitToFt;
                dy *= unitToFt;

                if (dx >= dy) { lengthFt = dx; widthFt = dy; }
                else { lengthFt = dy; widthFt = dx; }
            }
            catch
            {
                lengthFt = 0;
                widthFt = 0;
            }
        }
    }
}
