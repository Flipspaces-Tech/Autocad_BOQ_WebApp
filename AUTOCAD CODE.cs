// VisJsonNet.cs  (AutoCAD 2022 / .NET Framework, C# 7.3)
// References: acmgd.dll, acdbmgd.dll, accoremgd.dll  (Copy Local = False)
//
// ✅ Outputs:
//   1) vis_export_visibility.json        (same)
//   2) vis_export_all.json               (same)
//   3) vis_export_zone_rows.csv          (same, optional)
//   4) vis_export_sheet_like.json        (same)  <-- blocks (INSERTs)
//   5) vis_export_planner_dims.json      (same)  <-- planner dims
//   6) vis_export_layers_sheet_like.json (NEW)   <-- NON-BLOCK entities aggregated by (zone, layer)
//
// ✅ NEW RULE (your request):
//   - Any zone == "" OR "Unmarked Area"  -> "misc"
//   - Applies to BOTH sheet_like + layers outputs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.ApplicationServices.Core;
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
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null) doc.Editor.WriteMessage("\n✅ PINGNET OK\n");
        }

        [CommandMethod("VISJSONNET")]
        public void VisJsonNetCommand()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
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
            string outJsonSheetLike = Path.Combine(outDir, "vis_export_sheet_like.json");
            string outCsvZoneRows = Path.Combine(outDir, "vis_export_zone_rows.csv");
            string outJsonPlannerDims = Path.Combine(outDir, "vis_export_planner_dims.json");

            // ✅ NEW
            string outJsonLayersSheetLike = Path.Combine(outDir, "vis_export_layers_sheet_like.json");

            double unitToFt = GetScaleToFeet(db);

            var visGrouped = new Dictionary<string, VisAgg>(StringComparer.OrdinalIgnoreCase);
            var allGrouped = new Dictionary<string, AllAgg>(StringComparer.OrdinalIgnoreCase);

            int totalInserts = 0;
            int dynamicProcessed = 0;
            int visKept = 0;

            PlannerDims plannerDims = null;

            // ✅ NEW: non-block layer metrics aggregates (zone, layer)
            var layerAgg = new Dictionary<ZoneLayerKey, LayerAgg>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);

                // Build zones from PLANNER/planner
                var zones = CollectPlannerZones(tr, ms);

                // Create layer "planner" + move all entities from PLANNER → planner
                EnsureLayerExists(tr, db, "planner");
                MoveLayerContents(tr, ms, "PLANNER", "planner");

                // Compute planner dims from zone polygons
                plannerDims = ComputePlannerDims(zones, unitToFt);

                // ✅ NEW: compute NON-BLOCK entities layer metrics
                ComputeNonBlockLayerMetrics(tr, ms, zones, unitToFt, layerAgg);

                // ---------------- Existing INSERT loops ----------------
                foreach (ObjectId id in ms)
                {
                    if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                        continue;

                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                    if (br == null) continue;

                    // Skip planner content blocks refs (they define zones)
                    if (IsPlannerLayer(br.Layer))
                        continue;

                    totalInserts++;

                    string baseName = Safe(GetBaseName(tr, br));
                    if (string.IsNullOrWhiteSpace(baseName))
                        baseName = Safe(br.Name);

                    // size in feet
                    double lenFt, widFt;
                    GetBlockSizeFeet(tr, br, unitToFt, out lenFt, out widFt);

                    // distance params (Distance1..Distance4) in feet
                    double d1Ft, d2Ft, d3Ft, d4Ft;
                    GetDistanceParamsFeet(br, unitToFt, out d1Ft, out d2Ft, out d3Ft, out d4Ft);

                    // ✅ zone for this block ref
                    string zoneName = NormalizeZone(ResolveZoneForBlockRef(br, zones));

                    // dynamic/custom props for Config
                    Dictionary<string, string> dynConfig = CollectDynamicConfig(br);

                    // ---------- ALL aggregation ----------
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

                        // config props buckets
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

                        // zoneCounts
                        if (!string.IsNullOrWhiteSpace(zoneName))
                        {
                            int prev = 0;
                            a.ZoneCounts.TryGetValue(zoneName, out prev);
                            a.ZoneCounts[zoneName] = prev + 1;

                            // ✅ DON'T override the product/category just because it's zoned
                            if (string.IsNullOrWhiteSpace(a.CategoryLike))
                                a.CategoryLike = Safe(br.Layer);
                        }
                        else
                        {
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

                a.ConfigString = BuildConfigStringFromBuckets(a.ConfigBuckets);
            }

            // Write JSONs
            string jsonVis = BuildVisibilityJson(dwgFileName, totalInserts, dynamicProcessed, visKept, visGrouped.Values.ToList());
            File.WriteAllText(outJsonVisibility, jsonVis, Encoding.UTF8);

            string jsonAll = BuildAllJson(dwgFileName, totalInserts, allGrouped.Values.ToList());
            File.WriteAllText(outJsonAll, jsonAll, Encoding.UTF8);

            var zoneRows = BuildZoneRows(allGrouped.Values.ToList());
            WriteZoneRowsCsv(outCsvZoneRows, zoneRows);

            var sheetRows = BuildSheetLikeRows(allGrouped.Values.ToList());
            string sheetJson = BuildSheetLikeJson(dwgFileName, sheetRows);
            File.WriteAllText(outJsonSheetLike, sheetJson, Encoding.UTF8);

            string plannerJson = BuildPlannerDimsJson(dwgFileName, plannerDims);
            File.WriteAllText(outJsonPlannerDims, plannerJson, Encoding.UTF8);

            // ✅ NEW: build and write "layers" sheet json
            var layerRows = BuildLayersSheetLikeRows(layerAgg);
            string layersJson = BuildLayersSheetLikeJson(dwgFileName, "layers", layerRows);
            File.WriteAllText(outJsonLayersSheetLike, layersJson, Encoding.UTF8);

            doc.Editor.WriteMessage("\n✅ VISJSONNET wrote:\n");
            doc.Editor.WriteMessage("  - " + outJsonVisibility + "\n");
            doc.Editor.WriteMessage("  - " + outJsonAll + "\n");
            doc.Editor.WriteMessage("  - " + outCsvZoneRows + "\n");
            doc.Editor.WriteMessage("  - " + outJsonSheetLike + "   (sheet-like JSON with Config)\n");
            doc.Editor.WriteMessage("  - " + outJsonPlannerDims + "   (planner layer dims)\n");
            doc.Editor.WriteMessage("  - " + outJsonLayersSheetLike + "   (NON-BLOCK layer sheet → tab 'layers')\n");
        }

        // ======================= ✅ NEW: ZONE NORMALIZATION =======================
        private static string NormalizeZone(string zone)
        {
            string z = (zone ?? "").Trim();

            // your 2 cases:
            if (z.Length == 0) return "misc";
            if (z.Equals("Unmarked Area", StringComparison.OrdinalIgnoreCase)) return "misc";

            return z;
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

        private class SheetRow
        {
            public string Product;
            public string BoqName;
            public string Zone;
            public string Config;
            public int QtyValue;
            public double LengthFt;
            public double WidthFt;
            public string Params;
            public string Preview;
        }

        // ======================= PLANNER DIMS =======================

        private class PlannerDims
        {
            public double OverallWidthFt;
            public double OverallHeightFt;
            public readonly List<PlannerZoneDim> Zones = new List<PlannerZoneDim>();
        }

        private class PlannerZoneDim
        {
            public string Name;
            public double WidthFt;
            public double HeightFt;
            public double AreaSqFt;
        }

        // ======================= NEW: LAYERS SHEET MODELS =======================

        private struct ZoneLayerKey : IEquatable<ZoneLayerKey>
        {
            public string Zone;
            public string Layer;

            public ZoneLayerKey(string zone, string layer)
            {
                Zone = NormalizeZone(zone);
                Layer = (layer ?? "").Trim();
            }

            public bool Equals(ZoneLayerKey other)
                => string.Equals(Zone, other.Zone, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Layer, other.Layer, StringComparison.OrdinalIgnoreCase);

            public override bool Equals(object obj)
                => obj is ZoneLayerKey && Equals((ZoneLayerKey)obj);

            public override int GetHashCode()
            {
                unchecked
                {
                    int h1 = (Zone ?? "").ToUpperInvariant().GetHashCode();
                    int h2 = (Layer ?? "").ToUpperInvariant().GetHashCode();
                    return (h1 * 397) ^ h2;
                }
            }
        }

        private class LayerAgg
        {
            public string Zone;
            public string Layer;

            public double OpenLenFt;     // open length only
            public double PerimeterFt;   // closed boundary perimeter
            public double AreaFt2;       // closed area
        }

        private class LayerSheetRow
        {
            public string Layer;
            public string Zone;
            public string Remarks;
            public double LengthFt;
            public double WidthFt;
            public double PerimeterFt;
            public double AreaFt2;
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
                sb.Append("      \"zone\": ").Append(JsonStr(NormalizeZone(r.Zone))).Append(",\n");
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

        private static string BuildPlannerDimsJson(string dwg, PlannerDims d)
        {
            if (d == null) d = new PlannerDims();

            var sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"dwg\": ").Append(JsonStr(dwg)).Append(",\n");
            sb.Append("  \"plannerLayer\": ").Append(JsonStr("planner")).Append(",\n");
            sb.Append("  \"overall\": {\n");
            sb.Append("    \"width_ft\": ").Append(Num(d.OverallWidthFt)).Append(",\n");
            sb.Append("    \"height_ft\": ").Append(Num(d.OverallHeightFt)).Append("\n");
            sb.Append("  },\n");
            sb.Append("  \"zones\": [\n");

            var zs = d.Zones
                .Where(z => z != null)
                .OrderBy(z => z.Name ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            for (int i = 0; i < zs.Count; i++)
            {
                var z = zs[i];
                sb.Append("    {\n");
                sb.Append("      \"name\": ").Append(JsonStr(z.Name ?? "")).Append(",\n");
                sb.Append("      \"width_ft\": ").Append(Num(z.WidthFt)).Append(",\n");
                sb.Append("      \"height_ft\": ").Append(Num(z.HeightFt)).Append(",\n");
                sb.Append("      \"area_sqft\": ").Append(Num(z.AreaSqFt)).Append("\n");
                sb.Append("    }");
                if (i < zs.Count - 1) sb.Append(",");
                sb.Append("\n");
            }

            sb.Append("  ]\n");
            sb.Append("}\n");
            return sb.ToString();
        }

        // ✅ NEW: Layers sheet-like JSON
        private static string BuildLayersSheetLikeJson(string dwg, string tab, List<LayerSheetRow> rows)
        {
            if (rows == null) rows = new List<LayerSheetRow>();

            var sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"dwg\": ").Append(JsonStr(dwg)).Append(",\n");
            sb.Append("  \"tab\": ").Append(JsonStr(tab ?? "layers")).Append(",\n");
            sb.Append("  \"items\": [\n");

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.Append("    {\n");
                sb.Append("      \"layer\": ").Append(JsonStr(r.Layer ?? "")).Append(",\n");
                sb.Append("      \"zone\": ").Append(JsonStr(NormalizeZone(r.Zone))).Append(",\n");
                sb.Append("      \"remarks\": ").Append(JsonStr(r.Remarks ?? "")).Append(",\n");
                sb.Append("      \"length (ft)\": ").Append(Num(r.LengthFt)).Append(",\n");
                sb.Append("      \"width (ft)\": ").Append(Num(r.WidthFt)).Append(",\n");
                sb.Append("      \"perimeter\": ").Append(Num(r.PerimeterFt)).Append(",\n");
                sb.Append("      \"area (ft2)\": ").Append(Num(r.AreaFt2)).Append("\n");
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
                    string z = NormalizeZone(kv.Key ?? "");
                    int qty = kv.Value;
                    if (qty <= 0) continue;

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
                        Csv(NormalizeZone(r.Zone)) + "," +
                        r.Qty.ToString(CultureInfo.InvariantCulture) + "," +
                        Csv(r.CategoryLike)
                    );
                }
            }
        }

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
                        string zone = NormalizeZone(kv.Key ?? "");
                        int qty = kv.Value;
                        if (qty <= 0) continue;

                        rows.Add(new SheetRow
                        {
                            Product = PrettyProductName(it.CategoryLike),
                            BoqName = it.BlockName ?? "",
                            Zone = zone, // ✅ already normalized
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
                    // ✅ your requested default
                    rows.Add(new SheetRow
                    {
                        Product = PrettyProductName(it.CategoryLike),
                        BoqName = it.BlockName ?? "",
                        Zone = "misc",
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

                    string val = "";
                    if (p.Value != null)
                    {
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

        // ======================= PLANNER LAYER =======================

        private static void EnsureLayerExists(Transaction tr, Database db, string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return;

            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName)) return;

            lt.UpgradeOpen();
            var ltr = new LayerTableRecord { Name = layerName };
            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }

        private static void MoveLayerContents(Transaction tr, BlockTableRecord ms, string fromLayer, string toLayer)
        {
            if (ms == null) return;
            fromLayer = (fromLayer ?? "").Trim();
            toLayer = (toLayer ?? "").Trim();
            if (fromLayer.Length == 0 || toLayer.Length == 0) return;

            foreach (ObjectId id in ms)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Entity))))
                    continue;

                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (!string.Equals((ent.Layer ?? "").Trim(), fromLayer, StringComparison.OrdinalIgnoreCase))
                    continue;

                ent.UpgradeOpen();
                ent.Layer = toLayer;
            }
        }

        private static PlannerDims ComputePlannerDims(List<PlannerZone> zones, double unitToFt)
        {
            var d = new PlannerDims();

            if (zones == null || zones.Count == 0)
                return d;

            double minX = double.PositiveInfinity, minY = double.PositiveInfinity;
            double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;

            foreach (var z in zones)
            {
                if (z == null || z.Poly == null || z.Poly.Count < 3) continue;

                double zMinX = double.PositiveInfinity, zMinY = double.PositiveInfinity;
                double zMaxX = double.NegativeInfinity, zMaxY = double.NegativeInfinity;

                foreach (var p in z.Poly)
                {
                    if (p.X < zMinX) zMinX = p.X;
                    if (p.Y < zMinY) zMinY = p.Y;
                    if (p.X > zMaxX) zMaxX = p.X;
                    if (p.Y > zMaxY) zMaxY = p.Y;
                }

                double wFt = Math.Abs(zMaxX - zMinX) * unitToFt;
                double hFt = Math.Abs(zMaxY - zMinY) * unitToFt;

                double areaUnits2 = PolyArea(z.Poly);
                double areaFt2 = Math.Abs(areaUnits2) * unitToFt * unitToFt;

                d.Zones.Add(new PlannerZoneDim
                {
                    Name = z.Name ?? "",
                    WidthFt = wFt,
                    HeightFt = hFt,
                    AreaSqFt = areaFt2
                });

                if (zMinX < minX) minX = zMinX;
                if (zMinY < minY) minY = zMinY;
                if (zMaxX > maxX) maxX = zMaxX;
                if (zMaxY > maxY) maxY = zMaxY;
            }

            if (!double.IsInfinity(minX) && !double.IsInfinity(maxX))
            {
                d.OverallWidthFt = Math.Abs(maxX - minX) * unitToFt;
                d.OverallHeightFt = Math.Abs(maxY - minY) * unitToFt;
            }

            return d;
        }

        private static double PolyArea(List<Point2d> poly)
        {
            if (poly == null || poly.Count < 3) return 0;

            double area = 0;
            int n = poly.Count;
            for (int i = 0; i < n; i++)
            {
                var a = poly[i];
                var b = poly[(i + 1) % n];
                area += (a.X * b.Y - b.X * a.Y);
            }
            return 0.5 * area;
        }

        // ======================= ✅ NEW: NON-BLOCK → LAYERS METRICS =======================

        private static void ComputeNonBlockLayerMetrics(
            Transaction tr,
            BlockTableRecord ms,
            List<PlannerZone> zones,
            double unitToFt,
            Dictionary<ZoneLayerKey, LayerAgg> agg)
        {
            if (ms == null) return;
            if (agg == null) return;

            foreach (ObjectId id in ms)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Entity))))
                    continue;

                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                // skip block refs (user wants "everything not blocks")
                if (ent is BlockReference) continue;

                // skip planner layer entities
                if (IsPlannerLayer(ent.Layer)) continue;

                string layer = Safe(ent.Layer).Trim();
                if (layer.Length == 0) layer = "misc";

                // ✅ normalize zone here
                string zone = NormalizeZone(ResolveZoneForEntity(ent, zones));

                var key = new ZoneLayerKey(zone, layer);
                LayerAgg a;
                if (!agg.TryGetValue(key, out a))
                {
                    a = new LayerAgg { Zone = zone, Layer = layer, OpenLenFt = 0, PerimeterFt = 0, AreaFt2 = 0 };
                    agg[key] = a;
                }

                if (ent is Line)
                {
                    var ln = (Line)ent;
                    double lenFt = ln.Length * unitToFt;
                    if (lenFt > 0) a.OpenLenFt += lenFt;
                    continue;
                }

                if (ent is Arc)
                {
                    var ar = (Arc)ent;
                    double sweepRad = Math.Abs(ar.EndAngle - ar.StartAngle);
                    double lenFt = Math.Abs(ar.Radius * sweepRad) * unitToFt;
                    if (lenFt > 0) a.OpenLenFt += lenFt;
                    continue;
                }

                if (ent is Circle)
                {
                    var c = (Circle)ent;
                    double pFt = (2.0 * Math.PI * Math.Abs(c.Radius)) * unitToFt;
                    double aFt2 = (Math.PI * c.Radius * c.Radius) * unitToFt * unitToFt;
                    if (pFt > 0) a.PerimeterFt += pFt;
                    if (aFt2 > 0) a.AreaFt2 += aFt2;
                    continue;
                }

                if (ent is Polyline)
                {
                    var pl = (Polyline)ent;
                    double lenFt = pl.Length * unitToFt;

                    if (pl.Closed)
                    {
                        if (lenFt > 0) a.PerimeterFt += lenFt;
                        try
                        {
                            double areaFt2 = Math.Abs(pl.Area) * unitToFt * unitToFt;
                            if (areaFt2 > 0) a.AreaFt2 += areaFt2;
                        }
                        catch { }
                    }
                    else
                    {
                        if (lenFt > 0) a.OpenLenFt += lenFt;
                    }
                    continue;
                }

                if (ent is Polyline2d)
                {
                    var pl2 = (Polyline2d)ent;
                    try
                    {
                        double lenFt = GetCurveLengthFt(tr, pl2, unitToFt);
                        bool closed = pl2.Closed;
                        if (closed) a.PerimeterFt += lenFt;
                        else a.OpenLenFt += lenFt;
                    }
                    catch { }
                    continue;
                }

                if (ent is Polyline3d)
                {
                    var pl3 = (Polyline3d)ent;
                    try
                    {
                        double lenFt = GetCurveLengthFt(tr, pl3, unitToFt);
                        bool closed = pl3.Closed;
                        if (closed) a.PerimeterFt += lenFt;
                        else a.OpenLenFt += lenFt;
                    }
                    catch { }
                    continue;
                }

                if (ent is Hatch)
                {
                    var h = (Hatch)ent;
                    try
                    {
                        double areaFt2 = Math.Abs(h.Area) * unitToFt * unitToFt;
                        if (areaFt2 > 0) a.AreaFt2 += areaFt2;
                    }
                    catch { }
                    continue;
                }
            }
        }

        private static double GetCurveLengthFt(Transaction tr, Entity ent, double unitToFt)
        {
            try
            {
                Curve c = ent as Curve;
                if (c != null)
                {
                    double start = c.StartParam;
                    double end = c.EndParam;
                    double d0 = c.GetDistanceAtParameter(start);
                    double d1 = c.GetDistanceAtParameter(end);
                    return Math.Abs(d1 - d0) * unitToFt;
                }
            }
            catch { }
            return 0;
        }

        private static List<LayerSheetRow> BuildLayersSheetLikeRows(Dictionary<ZoneLayerKey, LayerAgg> agg)
        {
            var rows = new List<LayerSheetRow>();
            if (agg == null || agg.Count == 0) return rows;

            foreach (var kv in agg)
            {
                var a = kv.Value;
                if (a == null) continue;

                string zone = NormalizeZone(a.Zone ?? "");
                string layer = (a.Layer ?? "");

                if (a.OpenLenFt > 0)
                {
                    rows.Add(new LayerSheetRow
                    {
                        Layer = layer,
                        Zone = zone,
                        Remarks = "OPEN length only",
                        LengthFt = a.OpenLenFt,
                        WidthFt = 0,
                        PerimeterFt = 0,
                        AreaFt2 = 0
                    });
                }

                if (a.PerimeterFt > 0 || a.AreaFt2 > 0)
                {
                    double Lrec, Wrec;
                    SolveRectDimsFromPerimeterArea(a.PerimeterFt, a.AreaFt2, out Lrec, out Wrec);

                    rows.Add(new LayerSheetRow
                    {
                        Layer = layer,
                        Zone = zone,
                        Remarks = "CLOSED (rectangle): length/width + perimeter & area",
                        LengthFt = Lrec,
                        WidthFt = Wrec,
                        PerimeterFt = a.PerimeterFt,
                        AreaFt2 = a.AreaFt2
                    });
                }
            }

            return rows
                .OrderBy(r => r.Layer ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Zone ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Remarks ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void SolveRectDimsFromPerimeterArea(double P, double A, out double L, out double W)
        {
            L = 0; W = 0;
            try
            {
                if (P <= 0 || A <= 0) { L = 0; W = 0; return; }
                double S = P / 2.0;
                double D = S * S - 4.0 * A;
                if (D < -1e-9) { L = 0; W = 0; return; }
                if (D < 0) D = 0;
                double root = Math.Sqrt(D);
                double a = 0.5 * (S + root);
                double b = 0.5 * (S - root);
                if (a <= 0 || b <= 0) { L = 0; W = 0; return; }
                if (a >= b) { L = a; W = b; } else { L = b; W = a; }
            }
            catch { L = 0; W = 0; }
        }

        private static string ResolveZoneForEntity(Entity ent, List<PlannerZone> zones)
        {
            if (zones == null || zones.Count == 0 || ent == null) return "";

            try
            {
                Point2d pt;

                if (ent is Line)
                {
                    var ln = (Line)ent;
                    pt = new Point2d((ln.StartPoint.X + ln.EndPoint.X) * 0.5, (ln.StartPoint.Y + ln.EndPoint.Y) * 0.5);
                }
                else if (ent is Circle)
                {
                    var c = (Circle)ent;
                    pt = new Point2d(c.Center.X, c.Center.Y);
                }
                else if (ent is Arc)
                {
                    var a = (Arc)ent;
                    pt = new Point2d(a.Center.X, a.Center.Y);
                }
                else
                {
                    Extents3d? ext = TryGetGeometricExtents(ent);
                    if (ext == null) return "";
                    var e = ext.Value;
                    pt = new Point2d((e.MinPoint.X + e.MaxPoint.X) * 0.5, (e.MinPoint.Y + e.MaxPoint.Y) * 0.5);
                }

                foreach (var z in zones)
                {
                    if (z == null || z.Poly == null || z.Poly.Count < 3) continue;
                    if (PointInPolygon(pt, z.Poly))
                        return z.Name ?? "";
                }
            }
            catch { }

            return "";
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
                    case UnitsValue.Undefined: return 1.0;
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
        // PLANNER ZONES
        // =====================================================================

        private class PlannerZone
        {
            public string Name;
            public List<Point2d> Poly = new List<Point2d>();
        }

        private static bool IsPlannerLayer(string layer)
        {
            string s = (layer ?? "").Trim();
            if (s.Length == 0) return false;
            return s.Equals("PLANNER", StringComparison.OrdinalIgnoreCase)
                || s.Equals("planner", StringComparison.OrdinalIgnoreCase);
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

                if (!IsPlannerLayer(ent.Layer)) continue;

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