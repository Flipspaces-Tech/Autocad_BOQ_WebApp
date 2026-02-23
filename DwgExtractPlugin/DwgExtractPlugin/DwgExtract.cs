// VisJsonNet.cs  (AutoCAD 2022 / .NET Framework, C# 7.3)
// References: acmgd.dll, acdbmgd.dll, accoremgd.dll  (Copy Local = False)
//
// ✅ Outputs:
//   1) vis_export_visibility.json   (same as before)
//   2) vis_export_all.json          (same as before)
//   3) vis_export_zone_rows.csv     (same as before, optional)
//   4) ✅ vis_export_sheet_like.json  (NEW)  <-- for Google Sheets uploader
//
// “sheet-like” JSON schema:
// {
//   "dwg": "...",
//   "items": [
//      {
//        "Product": "...",
//        "BOQ name": "...",
//        "zone": "...",
//        "Config": "RADIUS=900; Visibility1=2",
//        "qty_value": 1,
//        "length (ft)": 2.4167,
//        "width (ft)": 0.1785,
//        "Params": "d1=1.97, d2=1.97",
//        "Preview": ""
//      }
//   ]
// }

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
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

            // ✅ NEW: “sheet-like” JSON
            string outJsonSheetLike = Path.Combine(outDir, "vis_export_sheet_like.json");

            // (optional) keep your old zone-rows file too
            string outCsvZoneRows = Path.Combine(outDir, "vis_export_zone_rows.csv");

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

                // ✅ Build zones from PLANNER (INSERT bboxes + closed polylines + text labels)
                var zones = CollectPlannerZones(tr, ms);

                foreach (ObjectId id in ms)
                {
                    if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                        continue;

                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                    if (br == null) continue;

                    // Skip the PLANNER block refs themselves (they define zones)
                    if (IsPlannerLayer(br.Layer))
                        continue;

                    totalInserts++;

                    string baseName = Safe(GetBaseName(tr, br));
                    if (string.IsNullOrWhiteSpace(baseName))
                        baseName = Safe(br.Name);

                    // ✅ size in feet
                    double lenFt, widFt;
                    GetBlockSizeFeet(tr, br, unitToFt, out lenFt, out widFt);

                    // ✅ distance params (Distance1..Distance4) in feet
                    double d1Ft, d2Ft, d3Ft, d4Ft;
                    GetDistanceParamsFeet(br, unitToFt, out d1Ft, out d2Ft, out d3Ft, out d4Ft);

                    // ✅ zone for this block ref
                    string zoneName = ResolveZoneForBlockRef(br, zones);

                    // ✅ Collect ALL dynamic/custom props for Config column (e.g. RADIUS, Visibility1)
                    //    We’ll aggregate them per block type.
                    Dictionary<string, string> dynConfig = CollectDynamicConfig(br);

                    // ---------- ALL aggregation (for default output + sheet-like) ----------
                    {
                        AllAgg a;
                        if (!allGrouped.TryGetValue(baseName, out a))
                        {
                            a = new AllAgg
                            {
                                BlockName = baseName,
                                TotalCount = 0,
                                DynamicCount = 0,
                                CategoryLike = ""
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

                        // ✅ store config props (for Config column in sheet-like)
                        if (dynConfig != null && dynConfig.Count > 0)
                        {
                            foreach (var kv in dynConfig)
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                                string k = kv.Key.Trim();
                                string v = (kv.Value ?? "").Trim();

                                List<string> bucket;
                                if (!a.ConfigBuckets.TryGetValue(k, out bucket))
                                {
                                    bucket = new List<string>();
                                    a.ConfigBuckets[k] = bucket;
                                }
                                if (!string.IsNullOrWhiteSpace(v))
                                    bucket.Add(v);
                            }
                        }

                        // ✅ zoneCounts
                        if (!string.IsNullOrWhiteSpace(zoneName))
                        {
                            int prev = 0;
                            a.ZoneCounts.TryGetValue(zoneName, out prev);
                            a.ZoneCounts[zoneName] = prev + 1;

                            // If zoned, categoryLike becomes PLANNER (like your earlier output)
                            a.CategoryLike = "PLANNER";
                        }
                        else
                        {
                            // fallback "product-like" = layer
                            if (string.IsNullOrWhiteSpace(a.CategoryLike))
                                a.CategoryLike = Safe(br.Layer);
                        }
                    }

                    // ---------- VISIBILITY aggregation ----------
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

            // finalize medians + config string
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

                // ✅ build one representative config string for this block type
                a.ConfigString = BuildConfigStringFromBuckets(a.ConfigBuckets);
            }

            // Write JSONs
            string jsonVis = BuildVisibilityJson(dwgFileName, totalInserts, dynamicProcessed, visKept, visGrouped.Values.ToList());
            File.WriteAllText(outJsonVisibility, jsonVis, Encoding.UTF8);

            string jsonAll = BuildAllJson(dwgFileName, totalInserts, allGrouped.Values.ToList());
            File.WriteAllText(outJsonAll, jsonAll, Encoding.UTF8);

            // (optional) zone rows CSV
            var zoneRows = BuildZoneRows(allGrouped.Values.ToList());
            WriteZoneRowsCsv(outCsvZoneRows, zoneRows);

            // ✅ NEW: sheet-like JSON for uploader (includes Config)
            var sheetRows = BuildSheetLikeRows(allGrouped.Values.ToList());
            string sheetJson = BuildSheetLikeJson(dwgFileName, sheetRows);
            File.WriteAllText(outJsonSheetLike, sheetJson, Encoding.UTF8);

            doc.Editor.WriteMessage("\n✅ VISJSONNET wrote:\n");
            doc.Editor.WriteMessage("  - " + outJsonVisibility + "\n");
            doc.Editor.WriteMessage("  - " + outJsonAll + "\n");
            doc.Editor.WriteMessage("  - " + outCsvZoneRows + "\n");
            doc.Editor.WriteMessage("  - " + outJsonSheetLike + "   (sheet-like JSON with Config)\n");
        }

        // ======================= MODELS =======================

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

            public readonly Dictionary<string, int> ZoneCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public string CategoryLike;

            // ✅ NEW: collect per-property values for Config column
            public readonly Dictionary<string, List<string>> ConfigBuckets = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            public string ConfigString = "";
        }

        private class ZoneRow
        {
            public string BlockName;
            public string Zone;
            public int Qty;
            public string CategoryLike;
        }

        // ✅ NEW: rows that directly match the sheet screenshot columns
        private class SheetRow
        {
            // We’ll output keys exactly as your sheet headers.
            public string Product;       // "Product"
            public string BoqName;       // "BOQ name"
            public string Zone;          // "zone"
            public string Config;        // "Config"
            public int QtyValue;         // "qty_value"
            public double LengthFt;      // "length (ft)"
            public double WidthFt;       // "width (ft)"
            public string Params;        // "Params"
            public string Preview;       // "Preview"
        }

        // ======================= JSON BUILDERS =======================

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
                sb.Append("      \"distance4_ft\": ").Append(Num(it.Distance4Ft)).Append(",\n");
                sb.Append("      \"categoryLike\": ").Append(JsonStr(it.CategoryLike ?? "")).Append(",\n");
                sb.Append("      \"zoneCounts\": ").Append(BuildZoneCountsJson(it.ZoneCounts)).Append("\n");
                sb.Append("    }");
                if (i < items.Count - 1) sb.Append(",");
                sb.Append("\n");
            }

            sb.Append("  ]\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        private static string BuildZoneCountsJson(Dictionary<string, int> map)
        {
            if (map == null || map.Count == 0) return "{}";

            var keys = map.Keys
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (keys.Count == 0) return "{}";

            var sb = new StringBuilder();
            sb.Append("{");
            for (int i = 0; i < keys.Count; i++)
            {
                string k = keys[i];
                int v = map[k];
                sb.Append(JsonStr(k)).Append(": ").Append(v.ToString(CultureInfo.InvariantCulture));
                if (i < keys.Count - 1) sb.Append(", ");
            }
            sb.Append("}");
            return sb.ToString();
        }

        // ✅ NEW: sheet-like JSON writer
        private static string BuildSheetLikeJson(string dwg, List<SheetRow> rows)
        {
            var sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"dwg\": ").Append(JsonStr(dwg)).Append(",\n");
            sb.Append("  \"items\": [\n");

            if (rows == null) rows = new List<SheetRow>();

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.Append("    {\n");
                sb.Append("      \"Product\": ").Append(JsonStr(r.Product ?? "")).Append(",\n");
                sb.Append("      \"BOQ name\": ").Append(JsonStr(r.BoqName ?? "")).Append(",\n");
                sb.Append("      \"zone\": ").Append(JsonStr(r.Zone ?? "")).Append(",\n");
                sb.Append("      \"Config\": ").Append(JsonStr(r.Config ?? "")).Append(",\n");
                sb.Append("      \"qty_value\": ").Append(r.QtyValue.ToString(CultureInfo.InvariantCulture)).Append(",\n");
                sb.Append("      \"length (ft)\": ").Append(Num(r.LengthFt)).Append(",\n");
                sb.Append("      \"width (ft)\": ").Append(Num(r.WidthFt)).Append(",\n");
                sb.Append("      \"Params\": ").Append(JsonStr(r.Params ?? "")).Append(",\n");
                sb.Append("      \"Preview\": ").Append(JsonStr(r.Preview ?? "")).Append("\n");
                sb.Append("    }");
                if (i < rows.Count - 1) sb.Append(",");
                sb.Append("\n");
            }

            sb.Append("  ]\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        // ======================= CSV BUILDERS (optional) =======================

        private static List<ZoneRow> BuildZoneRows(List<AllAgg> items)
        {
            var rows = new List<ZoneRow>();
            if (items == null) return rows;

            foreach (var it in items)
            {
                if (it == null || it.ZoneCounts == null || it.ZoneCounts.Count == 0)
                    continue;

                foreach (var kv in it.ZoneCounts)
                {
                    string z = kv.Key ?? "";
                    int qty = kv.Value;
                    if (string.IsNullOrWhiteSpace(z) || qty <= 0) continue;

                    rows.Add(new ZoneRow
                    {
                        BlockName = it.BlockName ?? "",
                        Zone = z,
                        Qty = qty,
                        CategoryLike = it.CategoryLike ?? ""
                    });
                }
            }

            return rows
                .OrderBy(r => r.BlockName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Zone, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void WriteZoneRowsCsv(string path, List<ZoneRow> rows)
        {
            using (var sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                sw.WriteLine("blockName,zone,qty,categoryLike");
                if (rows == null) return;

                foreach (var r in rows)
                {
                    sw.WriteLine(
                        Csv(r.BlockName) + "," +
                        Csv(r.Zone) + "," +
                        r.Qty.ToString(CultureInfo.InvariantCulture) + "," +
                        Csv(r.CategoryLike)
                    );
                }
            }
        }

        // ✅ Build rows shaped exactly like the sheet screenshot
        private static List<SheetRow> BuildSheetLikeRows(List<AllAgg> items)
        {
            var rows = new List<SheetRow>();
            if (items == null) return rows;

            foreach (var it in items)
            {
                if (it == null) continue;

                if (it.ZoneCounts != null && it.ZoneCounts.Count > 0)
                {
                    foreach (var kv in it.ZoneCounts)
                    {
                        string zone = kv.Key ?? "";
                        int qty = kv.Value;

                        if (string.IsNullOrWhiteSpace(zone) || qty <= 0) continue;

                        rows.Add(new SheetRow
                        {
                            Product = PrettyProductName(it.CategoryLike),
                            BoqName = it.BlockName ?? "",
                            Zone = zone,
                            Config = it.ConfigString ?? "",
                            QtyValue = qty,
                            LengthFt = it.LengthFt,
                            WidthFt = it.WidthFt,
                            Params = BuildParamsString(it),
                            Preview = ""
                        });
                    }
                }
                else
                {
                    rows.Add(new SheetRow
                    {
                        Product = PrettyProductName(it.CategoryLike),
                        BoqName = it.BlockName ?? "",
                        Zone = "Unmarked Area",
                        Config = it.ConfigString ?? "",
                        QtyValue = it.TotalCount,
                        LengthFt = it.LengthFt,
                        WidthFt = it.WidthFt,
                        Params = BuildParamsString(it),
                        Preview = ""
                    });
                }
            }

            return rows
                .OrderBy(r => r.Product, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.BoqName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Zone, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // If you want “Loose Furniture” instead of “PLANNER” etc, map here:
        private static string PrettyProductName(string categoryLike)
        {
            string s = (categoryLike ?? "").Trim();
            if (s.Length == 0) return "";

            if (s.Equals("PLANNER", StringComparison.OrdinalIgnoreCase))
                return "Loose Furniture";

            return s;
        }

        private static string BuildParamsString(AllAgg it)
        {
            var parts = new List<string>();

            if (it.Distance1Ft > 0) parts.Add("d1=" + it.Distance1Ft.ToString("0.##", CultureInfo.InvariantCulture));
            if (it.Distance2Ft > 0) parts.Add("d2=" + it.Distance2Ft.ToString("0.##", CultureInfo.InvariantCulture));
            if (it.Distance3Ft > 0) parts.Add("d3=" + it.Distance3Ft.ToString("0.##", CultureInfo.InvariantCulture));
            if (it.Distance4Ft > 0) parts.Add("d4=" + it.Distance4Ft.ToString("0.##", CultureInfo.InvariantCulture));

            if (parts.Count == 0) return "";
            return string.Join(", ", parts);
        }

        // ======================= CONFIG (Dynamic/Custom props) =======================

        // ✅ Collect all dynamic properties into key/value (RADIUS, Visibility1, etc.)
        private static Dictionary<string, string> CollectDynamicConfig(BlockReference br)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                if (br == null) return map;
                if (!br.IsDynamicBlock) return map;

                var props = br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty p in props)
                {
                    if (p == null) continue;

                    string name = (p.PropertyName ?? "").Trim();
                    if (name.Length == 0) continue;

                    // Skip very noisy/system-like ones if they appear in your files
                    // (Keep visibility + radius + custom values)
                    // You can add more filters if needed.
                    // if (name.Equals("Origin", StringComparison.OrdinalIgnoreCase)) continue;

                    string val = "";
                    if (p.Value != null)
                    {
                        // normalize numeric values nicely
                        double dv;
                        if (TryGetDouble(p.Value, out dv))
                            val = dv.ToString("0.####", CultureInfo.InvariantCulture);
                        else
                            val = p.Value.ToString();
                    }

                    if (!string.IsNullOrWhiteSpace(val))
                        map[name] = val;
                }
            }
            catch { }

            return map;
        }

        // ✅ Turn buckets into one string: "RADIUS=900; Visibility1=2"
        // Rule:
        // - if numeric -> choose median
        // - else -> choose most common value (mode)
        private static string BuildConfigStringFromBuckets(Dictionary<string, List<string>> buckets)
        {
            if (buckets == null || buckets.Count == 0) return "";

            var keys = buckets.Keys
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var parts = new List<string>();

            foreach (var k in keys)
            {
                List<string> vals;
                if (!buckets.TryGetValue(k, out vals) || vals == null || vals.Count == 0)
                    continue;

                // try numeric median
                var nums = new List<double>();
                foreach (var s in vals)
                {
                    double d;
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                        nums.Add(d);
                }

                string chosen;
                if (nums.Count >= Math.Max(1, vals.Count / 2))
                {
                    nums.Sort();
                    chosen = (nums.Count % 2 == 1)
                        ? nums[nums.Count / 2].ToString("0.####", CultureInfo.InvariantCulture)
                        : (0.5 * (nums[nums.Count / 2 - 1] + nums[nums.Count / 2])).ToString("0.####", CultureInfo.InvariantCulture);
                }
                else
                {
                    // mode for strings
                    chosen = vals
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .GroupBy(x => x.Trim(), StringComparer.OrdinalIgnoreCase)
                        .OrderByDescending(g => g.Count())
                        .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(g => g.Key)
                        .FirstOrDefault() ?? "";
                }

                if (!string.IsNullOrWhiteSpace(chosen))
                    parts.Add(k + "=" + chosen);
            }

            if (parts.Count == 0) return "";
            return string.Join("; ", parts);
        }

        // ======================= HELPERS =======================

        private static string Csv(string s)
        {
            if (s == null) s = "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
            {
                s = s.Replace("\"", "\"\"");
                return "\"" + s + "\"";
            }
            return s;
        }

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
                        return 1.0;
                }
                return 1.0;
            }
            catch { return 1.0; }
        }

        // ======================= BLOCK NAME / DYNAMIC =======================

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

        // ======================= DISTANCE PARAMS =======================

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

                    int idx = ParseTrailingIndex(up);
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

        // ======================= SIZE (FT) =======================

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

        // =====================================================================
        // ✅ PLANNER ZONES
        // =====================================================================

        private class PlannerZone
        {
            public string Name;
            public List<Point2d> Poly = new List<Point2d>();
        }

        private static bool IsPlannerLayer(string layer)
        {
            return string.Equals((layer ?? "").Trim(), "PLANNER", StringComparison.OrdinalIgnoreCase);
        }

        private static List<PlannerZone> CollectPlannerZones(Transaction tr, BlockTableRecord ms)
        {
            var zones = new List<PlannerZone>();
            var labels = new List<Tuple<string, Point2d>>();

            foreach (ObjectId id in ms)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Entity))))
                    continue;

                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent is DBText)
                {
                    var t = (DBText)ent;
                    string txt = (t.TextString ?? "").Trim();
                    if (txt.Length > 0)
                        labels.Add(Tuple.Create(txt, new Point2d(t.Position.X, t.Position.Y)));
                }
                else if (ent is MText)
                {
                    var mt = (MText)ent;
                    string raw = (mt.Contents ?? "").Trim();
                    if (raw.Length > 0)
                    {
                        string first = raw;
                        int nl = first.IndexOf("\\P", StringComparison.OrdinalIgnoreCase);
                        if (nl >= 0) first = first.Substring(0, nl);
                        nl = first.IndexOf("\n", StringComparison.OrdinalIgnoreCase);
                        if (nl >= 0) first = first.Substring(0, nl);
                        first = first.Trim();
                        if (first.Length > 0)
                            labels.Add(Tuple.Create(first, new Point2d(mt.Location.X, mt.Location.Y)));
                    }
                }
            }

            // A) PLANNER INSERT zones (bbox rect)
            foreach (ObjectId id in ms)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                    continue;

                var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                if (br == null) continue;
                if (!IsPlannerLayer(br.Layer)) continue;

                Extents3d? ext = TryGetGeometricExtents(br);
                if (ext == null) continue;

                var e = ext.Value;
                var poly = new List<Point2d>
                {
                    new Point2d(e.MinPoint.X, e.MinPoint.Y),
                    new Point2d(e.MaxPoint.X, e.MinPoint.Y),
                    new Point2d(e.MaxPoint.X, e.MaxPoint.Y),
                    new Point2d(e.MinPoint.X, e.MaxPoint.Y),
                };

                string name = TryGetPlannerNameFromAttributes(tr, br);
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = Safe(GetBaseName(tr, br));
                    if (string.IsNullOrWhiteSpace(name)) name = Safe(br.Name);
                }

                zones.Add(new PlannerZone { Name = name.Trim(), Poly = poly });
            }

            // B) PLANNER closed polylines
            var polylineZones = new List<PlannerZone>();
            foreach (ObjectId id in ms)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Entity))))
                    continue;

                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!IsPlannerLayer(ent.Layer)) continue;

                if (ent is Polyline)
                {
                    var pl = (Polyline)ent;
                    if (!pl.Closed) continue;

                    var pts = new List<Point2d>();
                    int n = pl.NumberOfVertices;
                    if (n < 3) continue;

                    for (int i = 0; i < n; i++)
                    {
                        var p = pl.GetPoint2dAt(i);
                        pts.Add(new Point2d(p.X, p.Y));
                    }

                    pts = NormalizePoly(pts);
                    if (pts.Count >= 3)
                        polylineZones.Add(new PlannerZone { Name = "", Poly = pts });
                }
            }

            // Name polylines from labels (inside first, else nearest to centroid)
            var usedLabel = new HashSet<int>();
            int autoIndex = 1;

            foreach (var z in polylineZones)
            {
                string zname = "";

                for (int i = 0; i < labels.Count; i++)
                {
                    if (usedLabel.Contains(i)) continue;
                    var txt = labels[i].Item1;
                    var pt = labels[i].Item2;
                    if (string.IsNullOrWhiteSpace(txt)) continue;

                    if (PointInPolygon(pt, z.Poly))
                    {
                        zname = txt.Trim();
                        usedLabel.Add(i);
                        break;
                    }
                }

                if (string.IsNullOrWhiteSpace(zname) && labels.Count > 0)
                {
                    var c = PolyCentroid(z.Poly);
                    double best = double.MaxValue;
                    int bestIdx = -1;

                    for (int i = 0; i < labels.Count; i++)
                    {
                        if (usedLabel.Contains(i)) continue;
                        var txt = labels[i].Item1;
                        if (string.IsNullOrWhiteSpace(txt)) continue;

                        var p = labels[i].Item2;
                        double dx = p.X - c.X;
                        double dy = p.Y - c.Y;
                        double d2 = dx * dx + dy * dy;
                        if (d2 < best)
                        {
                            best = d2;
                            bestIdx = i;
                        }
                    }

                    if (bestIdx >= 0)
                    {
                        zname = labels[bestIdx].Item1.Trim();
                        usedLabel.Add(bestIdx);
                    }
                }

                if (string.IsNullOrWhiteSpace(zname))
                    zname = "Zone " + autoIndex.ToString("00", CultureInfo.InvariantCulture);

                autoIndex++;
                z.Name = zname;
                zones.Add(z);
            }

            // de-dupe
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var outZones = new List<PlannerZone>();

            foreach (var z in zones)
            {
                if (z == null || z.Poly == null || z.Poly.Count < 3) continue;
                string sig = z.Name + "|" + string.Join(";", z.Poly.Select(p => p.X.ToString("0.###", CultureInfo.InvariantCulture) + "," + p.Y.ToString("0.###", CultureInfo.InvariantCulture)));
                if (seen.Add(sig))
                    outZones.Add(z);
            }

            return outZones;
        }

        private static string TryGetPlannerNameFromAttributes(Transaction tr, BlockReference br)
        {
            try
            {
                if (br.AttributeCollection == null) return "";

                var cand = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "NAME", "ROOM", "ZONE", "LABEL", "TITLE" };

                foreach (ObjectId aid in br.AttributeCollection)
                {
                    if (aid.IsNull) continue;
                    var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                    if (ar == null) continue;

                    string tag = (ar.Tag ?? "").Trim();
                    if (tag.Length == 0) continue;

                    if (cand.Contains(tag))
                    {
                        string txt = (ar.TextString ?? "").Trim();
                        if (txt.Length > 0) return txt;
                    }
                }
            }
            catch { }
            return "";
        }

        private static string ResolveZoneForBlockRef(BlockReference br, List<PlannerZone> zones)
        {
            if (zones == null || zones.Count == 0) return "";

            try
            {
                Point3d p3 = br.Position;
                var p = new Point2d(p3.X, p3.Y);
                string z = ZoneForPoint(p, zones);
                if (!string.IsNullOrWhiteSpace(z)) return z;
            }
            catch { }

            try
            {
                Extents3d? ext = TryGetGeometricExtents(br);
                if (ext != null)
                {
                    var e = ext.Value;
                    var c = new Point2d(
                        (e.MinPoint.X + e.MaxPoint.X) * 0.5,
                        (e.MinPoint.Y + e.MaxPoint.Y) * 0.5
                    );
                    string z = ZoneForPoint(c, zones);
                    if (!string.IsNullOrWhiteSpace(z)) return z;
                }
            }
            catch { }

            return "";
        }

        private static string ZoneForPoint(Point2d pt, List<PlannerZone> zones)
        {
            foreach (var z in zones)
            {
                if (z == null || z.Poly == null || z.Poly.Count < 3) continue;
                if (PointInPolygon(pt, z.Poly))
                    return z.Name ?? "";
            }
            return "";
        }

        private static Extents3d? TryGetGeometricExtents(Entity ent)
        {
            try
            {
                Extents3d e = ent.GeometricExtents;
                if (double.IsNaN(e.MinPoint.X) || double.IsNaN(e.MaxPoint.X)) return null;
                return e;
            }
            catch { return null; }
        }

        private static List<Point2d> NormalizePoly(List<Point2d> poly)
        {
            if (poly == null || poly.Count == 0) return poly;

            var cleaned = new List<Point2d>();
            for (int i = 0; i < poly.Count; i++)
            {
                var p = poly[i];
                if (cleaned.Count == 0)
                {
                    cleaned.Add(p);
                }
                else
                {
                    var prev = cleaned[cleaned.Count - 1];
                    if (Math.Abs(p.X - prev.X) > 1e-9 || Math.Abs(p.Y - prev.Y) > 1e-9)
                        cleaned.Add(p);
                }
            }

            if (cleaned.Count > 2)
            {
                var a = cleaned[0];
                var b = cleaned[cleaned.Count - 1];
                if (Math.Abs(a.X - b.X) < 1e-9 && Math.Abs(a.Y - b.Y) < 1e-9)
                    cleaned.RemoveAt(cleaned.Count - 1);
            }

            return cleaned;
        }

        private static Point2d PolyCentroid(List<Point2d> poly)
        {
            if (poly == null || poly.Count == 0) return new Point2d(0, 0);
            double sx = 0, sy = 0;
            for (int i = 0; i < poly.Count; i++)
            {
                sx += poly[i].X;
                sy += poly[i].Y;
            }
            return new Point2d(sx / poly.Count, sy / poly.Count);
        }

        private static bool PointInPolygon(Point2d pt, List<Point2d> poly)
        {
            double x = pt.X, y = pt.Y;
            bool inside = false;
            int n = poly.Count;
            if (n < 3) return false;

            for (int i = 0; i < n; i++)
            {
                var a = poly[i];
                var b = poly[(i + 1) % n];

                bool cond = ((a.Y > y) != (b.Y > y));
                if (cond)
                {
                    double t = (y - a.Y) / (b.Y - a.Y + 1e-12);
                    double xInt = a.X + (b.X - a.X) * t;
                    if (x < xInt) inside = !inside;
                }
            }
            return inside;
        }
    }
}
