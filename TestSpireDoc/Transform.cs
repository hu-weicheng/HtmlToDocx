using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;    //AltChunk
using Spire.Doc;
using Spire.Doc.Interface;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace TestSpireDoc
{
    internal class Transform
    {
        Spire.Doc.Document doc = new Spire.Doc.Document();
        public string inputHtmlPath = string.Empty;
        public string outputDocxPath = string.Empty;
        public Transform() 
        {
            this.inputHtmlPath = @"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm";
            this.outputDocxPath = @"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx";
        }
        ~Transform()
        {
            if(this.doc!=null)
            {
                doc.Close();
            }
        }
        public bool Check(string inputHtmlPath, string outputHtmlPath)
        {
            return System.IO.File.Exists(this.inputHtmlPath)& System.IO.File.Exists(this.outputDocxPath);
            //return false;
        }
        public bool Translate()
        {
            bool exists = System.IO.File.Exists(this.inputHtmlPath);
            //Console.WriteLine(exists);
            if(exists)
            {
                // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
                doc.LoadFromFile(this.inputHtmlPath,
                                 FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
                //Console.WriteLine(doc.ToString());
                doc.SaveToFile(this.outputDocxPath, FileFormat.Docx);
                //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            }

            return exists;
        }

        public static void TestTranslate(string inputHtmlFilePath,string outputDocxFilePath)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            bool exists = System.IO.File.Exists(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm");
            //Console.WriteLine(exists);
            // XHTMLValidationType.None 可容忍普通 HTML 的小瑕疵
            doc.LoadFromFile(@"C:\Users\Raise\Downloads\6\2601U39467E-SA(OD1806)India\GSM 900 Head Left Cheek Mid-2\GSM 900 Head Left Cheek Mid-2.htm",
                             FileFormat.Html, Spire.Doc.Documents.XHTMLValidationType.None);
            //Console.WriteLine(doc.ToString());
            doc.SaveToFile(@"C:\Users\Raise\Desktop\TestSpire\GSM900_Report.docx", FileFormat.Docx);
            //doc.SaveToFile(@"C:\Reports\SAR_Report.docx", FileFormat.Docx2016);
            doc.Close();
        }
    }

    internal class MergeDocx
    {

        public static void MergeDocsEmbedded(string[] docxFiles, string outputPath)
        {
            File.Copy(docxFiles[0], outputPath, true);

            using (WordprocessingDocument mainDoc =
                WordprocessingDocument.Open(outputPath, true))
            {
                MainDocumentPart mainPart = mainDoc.MainDocumentPart;
                DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                foreach (var file in docxFiles.Skip(1))
                {
                    // 把待合并 docx 作为部件嵌入主文档
                    AlternativeFormatImportPart chunk =
                        mainPart.AddAlternativeFormatImportPart(
                            AlternativeFormatImportPartType.WordprocessingML);

                    using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        chunk.FeedData(fs);
                    }

                    //页尾加分割符
                    body.Append(new Paragraph(
                            new Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page })
                        ));

                    // ✅ 终极兜底：通过 MainDocumentPart 查 ID
                    string id = mainPart.GetIdOfPart(chunk);
                    body.Append(new AltChunk { Id = id });
                }

                mainPart.Document.Save();
            }
        }

        //页尾加分割符
        //public static void splitPage()
        //{
        //    body.Append(new Paragraph(
        //            new Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page })
        //        ));
        //    body.Append(new AltChunk { Id = chunk.Id });
        //}
    }


    //1.顺序2G 3G 4G 5G Wifi BT
    //2.它名称里不包含2G 3G 4G 5G Wifi BT
    // 2G的名称可能GSM 850、GSM 900、GSM 1800、GSM 1900
    // 2G的名称可能GSM 850、GSM 900、DCS 1800、PCS 1900
    //    3G的名称WCDMA
    // 4G的名称LTE
    // 5G的名称NR
    //3.顺序Head 、Body 、Limbs
    //4.顺序头部，Left 、 Right
    //5.先Cheek再Tilt
    //7.信道顺序，Low Mid High
    //8.Body的顺序，Front Back Left Right Top Bottom
    //9.Limbs的顺序Front Back Left Right Top Bottom
    //10.3G 4G 5G按Band的顺序排
    //11.5G的band名称n1 n2 n3...
    //12.2G按band的频率排，850 900 1800 1900
    //13.排序优先级，网络类型 身体部位，一个频段的顺序连起来
    //14.4G和5G的RB排序，从小到大，如 1RB 50%RB 100%RB
    public static class SarFileSorter
    {

        public static string[] DocxAsortV4(string[] docxFiles)
        {
            if (docxFiles == null || docxFiles.Length == 0)
                return docxFiles;

            /* ========== 1️⃣ 网络类型 ========== */
            var networkOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["GSM"] = 1,
                ["DCS"] = 1,
                ["PCS"] = 1,
                ["WCDMA"] = 2,
                ["LTE"] = 3,
                ["NR"] = 4,
                ["Wifi"] = 5,
                ["BT"] = 6
            };

            /* ========== 2️⃣ 身体部位 ========== */
            var regionOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Head"] = 1,
                ["Body"] = 2,
                ["Limbs"] = 3
            };

            /* ========== 3️⃣ 子部位 ========== */
            var subPartOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Front"] = 1,
                ["Back"] = 2,
                ["Left"] = 3,
                ["Right"] = 4,
                ["Top"] = 5,
                ["Bottom"] = 6
            };

            /* ========== 4️⃣ Head 左右 ========== */
            var headSideOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Left"] = 1,
                ["Right"] = 2
            };

            /* ========== 5️⃣ Cheek / Tilt ========== */
            var poseOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Cheek"] = 1,
                ["Tilt"] = 2
            };

            /* ========== 6️⃣ 信道 ========== */
            var channelOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Low"] = 1,
                ["Mid"] = 2,
                ["High"] = 3
            };

            /* ========== 工具函数 ========== */
            int MatchKey(string name, Dictionary<string, int> map)
            {
                foreach (var k in map.Keys)
                    if (name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                        return map[k];
                return int.MaxValue;
            }

            bool ContainsAny(string text, params string[] patterns)
            {
                foreach (var p in patterns)
                    if (text.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                return false;
            }

            /* ========== 2G 频率 ========== */
            int Get2GFreq(string name)
            {
                if (ContainsAny(name, "GSM 850", "GSM850")) return 850;
                if (ContainsAny(name, "GSM 900", "GSM900")) return 900;
                if (ContainsAny(name, "DCS 1800", "DCS1800", "GSM 1800", "GSM1800")) return 1800;
                if (ContainsAny(name, "PCS 1900", "PCS1900", "GSM 1900", "GSM1900")) return 1900;

                var m = Regex.Match(name, @"(850|900|1800|1900)");
                if (m.Success)
                {
                    switch (m.Value)
                    {
                        case "850": return 850;
                        case "900": return 900;
                        case "1800": return 1800;
                        case "1900": return 1900;
                        default: return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }

            /* ========== Band ========== */
            int GetBandNumber(string name)
            {
                var m = Regex.Match(name, @"(?:Band|B)\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            /* ========== 5G n1/n2/n3 ========== */
            int Get5GBand(string name)
            {
                var m = Regex.Match(name, @"n\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            /* ========== 第 14 条：RB 排序 ========== */
            int GetRbValue(string name)
            {
                var m = Regex.Match(name, @"(\d+)\s*%?\s*RB", RegexOptions.IgnoreCase);
                if (m.Success)
                    return int.Parse(m.Groups[1].Value);

                return int.MaxValue;
            }

            /* ========== 主排序 ========== */
            return docxFiles
                .Select(f =>
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    var network = MatchKey(name, networkOrder);

                    return new
                    {
                        File = f,
                        Network = network,
                        Region = MatchKey(name, regionOrder),
                        SubPart = MatchKey(name, subPartOrder),
                        HeadSide = MatchKey(name, headSideOrder),
                        Pose = MatchKey(name, poseOrder),
                        Channel = MatchKey(name, channelOrder),

                        FreqOrBand = network switch
                        {
                            1 => Get2GFreq(name),
                            2 => GetBandNumber(name),
                            3 => GetBandNumber(name),
                            4 => Get5GBand(name),
                            _ => int.MaxValue
                        },

                        Rb = GetRbValue(name)
                    };
                })
                .OrderBy(x => x.Network)      // ① 网络类型
                .ThenBy(x => x.FreqOrBand)    // ② 同一频段连续
                .ThenBy(x => x.Region)        // ③ 身体部位
                .ThenBy(x => x.SubPart)       // ④ 子部位
                .ThenBy(x => x.HeadSide)
                .ThenBy(x => x.Pose)
                .ThenBy(x => x.Rb)            // ⑤ 第 14 条：RB 从小到大
                .ThenBy(x => x.Channel)
                .Select(x => x.File)
                .ToArray();
        }
        public static string[] DocxAsortV5(string[] docxFiles)
        {
            if (docxFiles == null || docxFiles.Length == 0)
                return docxFiles;

            var networkOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["GSM"] = 1,
                ["DCS"] = 1,
                ["PCS"] = 1,
                ["WCDMA"] = 2,
                ["LTE"] = 3,
                ["NR"] = 4,
                ["Wifi"] = 5,
                ["BT"] = 6
            };

            var regionOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Head"] = 1,
                ["Body"] = 2,
                ["Limbs"] = 3
            };

            var subPartOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Front"] = 1,
                ["Back"] = 2,
                ["Left"] = 3,
                ["Right"] = 4,
                ["Top"] = 5,
                ["Bottom"] = 6
            };

            var headSideOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Left"] = 1,
                ["Right"] = 2
            };

            var poseOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Cheek"] = 1,
                ["Tilt"] = 2
            };

            var channelOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Low"] = 1,
                ["Mid"] = 2,
                ["High"] = 3
            };

            var RbOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["1RB"] = 1,
                ["50%RB"] = 2,
                ["50%"] = 2,
                ["50"] = 2,
                ["100"] = 3,
                ["100%"] = 3,
                ["100%RB"] = 3
            };

            int MatchKey(string name, Dictionary<string, int> map)
            {
                foreach (var k in map.Keys)
                    if (name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                        return map[k];
                return int.MaxValue;
            }

            bool ContainsAny(string text, params string[] patterns)
            {
                foreach (var p in patterns)
                    if (text.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                return false;
            }

            int Get2GFreq(string name)
            {
                if (ContainsAny(name, "GSM 850", "GSM850")) return 850;
                if (ContainsAny(name, "GSM 900", "GSM900")) return 900;
                if (ContainsAny(name, "DCS 1800", "DCS1800", "GSM 1800", "GSM1800")) return 1800;
                if (ContainsAny(name, "PCS 1900", "PCS1900", "GSM 1900", "GSM1900")) return 1900;

                var m = Regex.Match(name, @"(850|900|1800|1900)");
                if (m.Success)
                {
                    switch (m.Value)
                    {
                        case "850": return 850;
                        case "900": return 900;
                        case "1800": return 1800;
                        case "1900": return 1900;
                        default: return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }

            int GetBandNumber(string name)
            {
                var m = Regex.Match(name, @"(?:Band|B)\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            int Get5GBand(string name)
            {
                var m = Regex.Match(name, @"n\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            int GetRbValue(string name)
            {
                var m = Regex.Match(name, @"(\d+)\s*%?\s*RB", RegexOptions.IgnoreCase);
                if (m.Success)
                    return int.Parse(m.Groups[1].Value);
                return int.MaxValue;
            }

            return docxFiles
                .Select(f =>
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    var network = MatchKey(name, networkOrder);

                    return new
                    {
                        File = f,
                        Network = network,
                        Region = MatchKey(name, regionOrder),
                        SubPart = MatchKey(name, subPartOrder),
                        HeadSide = MatchKey(name, headSideOrder),
                        Pose = MatchKey(name, poseOrder),
                        Channel = MatchKey(name, channelOrder),
                        Rb = GetRbValue(name),

                        FreqOrBand = network switch
                        {
                            1 => Get2GFreq(name),
                            2 => GetBandNumber(name),
                            3 => GetBandNumber(name),
                            4 => Get5GBand(name),
                            _ => int.MaxValue
                        }
                    };
                })
                .OrderBy(x => x.Network)
                .ThenBy(x => x.FreqOrBand)
                .ThenBy(x => x.Region)
                .ThenBy(x => x.SubPart)
                .ThenBy(x => x.HeadSide)
                .ThenBy(x => x.Pose)
                .ThenBy(x => x.Channel)      // ✅ 信道优先
                .ThenBy(x => x.Rb)           // ✅ 同一信道内按 RB 排
                .Select(x => x.File)
                .ToArray();
        }
        public static string[] DocxAsortV6(string[] docxFiles)
        {
            if (docxFiles == null || docxFiles.Length == 0)
                return docxFiles;

            var networkOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["GSM"] = 1,
                ["DCS"] = 1,
                ["PCS"] = 1,
                ["WCDMA"] = 2,
                ["LTE"] = 3,
                ["NR"] = 4,
                ["Wifi"] = 5,
                ["BT"] = 6
            };

            var regionOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Head"] = 1,
                ["Body"] = 2,
                ["Limbs"] = 3
            };

            var subPartOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Front"] = 1,
                ["Back"] = 2,
                ["Left"] = 3,
                ["Right"] = 4,
                ["Top"] = 5,
                ["Bottom"] = 6
            };

            var headSideOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Left"] = 1,
                ["Right"] = 2
            };

            var poseOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Cheek"] = 1,
                ["Tilt"] = 2
            };

            var channelOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Low"] = 1,
                ["Mid"] = 2,
                ["High"] = 3
            };

            // ✅ RB 排序字典（枚举型排序）
            var RbOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["1RB"] = 1,
                ["50%RB"] = 2,
                ["50%"] = 2,
                ["50"] = 2,
                ["100"] = 3,
                ["100%"] = 3,
                ["100%RB"] = 3
            };

            int MatchKey(string name, Dictionary<string, int> map)
            {
                foreach (var k in map.Keys)
                    if (name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                        return map[k];
                return int.MaxValue;
            }

            bool ContainsAny(string text, params string[] patterns)
            {
                foreach (var p in patterns)
                    if (text.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                return false;
            }

            int Get2GFreq(string name)
            {
                if (ContainsAny(name, "GSM 850", "GSM850")) return 850;
                if (ContainsAny(name, "GSM 900", "GSM900")) return 900;
                if (ContainsAny(name, "DCS 1800", "DCS1800", "GSM 1800", "GSM1800")) return 1800;
                if (ContainsAny(name, "PCS 1900", "PCS1900", "GSM 1900", "GSM1900")) return 1900;

                var m = Regex.Match(name, @"(850|900|1800|1900)");
                if (m.Success)
                {
                    switch (m.Value)
                    {
                        case "850": return 850;
                        case "900": return 900;
                        case "1800": return 1800;
                        case "1900": return 1900;
                        default: return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }

            int GetBandNumber(string name)
            {
                var m = Regex.Match(name, @"(?:Band|B)\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            int Get5GBand(string name)
            {
                var m = Regex.Match(name, @"n\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            // ✅ 使用 RB 字典排序（关键）
            int GetRbOrder(string name)
            {
                foreach (var kv in RbOrder)
                {
                    if (name.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        return kv.Value;
                }
                return int.MaxValue;
            }

            return docxFiles
                .Select(f =>
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    var network = MatchKey(name, networkOrder);

                    return new
                    {
                        File = f,
                        Network = network,
                        Region = MatchKey(name, regionOrder),
                        SubPart = MatchKey(name, subPartOrder),
                        HeadSide = MatchKey(name, headSideOrder),
                        Pose = MatchKey(name, poseOrder),
                        Channel = MatchKey(name, channelOrder),
                        Rb = GetRbOrder(name),   // ✅ 用字典排序

                        FreqOrBand = network switch
                        {
                            1 => Get2GFreq(name),
                            2 => GetBandNumber(name),
                            3 => GetBandNumber(name),
                            4 => Get5GBand(name),
                            _ => int.MaxValue
                        }
                    };
                })
                .OrderBy(x => x.Network)
                .ThenBy(x => x.FreqOrBand)
                .ThenBy(x => x.Region)
                .ThenBy(x => x.SubPart)
                .ThenBy(x => x.HeadSide)
                .ThenBy(x => x.Pose)
                .ThenBy(x => x.Channel)   // ✅ 信道优先
                .ThenBy(x => x.Rb)        // ✅ RB 排序生效
                .Select(x => x.File)
                .ToArray();
        }
        public static string[] DocxAsortV7(string[] docxFiles)
        {
            if (docxFiles == null || docxFiles.Length == 0)
                return docxFiles;

            var networkOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["GSM"] = 1,
                ["DCS"] = 1,
                ["PCS"] = 1,
                ["WCDMA"] = 2,
                ["LTE"] = 3,
                ["NR"] = 4,
                ["Wifi"] = 5,
                ["WLAN"] = 5,
                ["BT"] = 6,
                ["Bluetooth"] = 6
            };

            var regionOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Head"] = 1,
                ["Body"] = 2,
                ["Limbs"] = 3
            };

            var subPartOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Front"] = 1,
                ["Back"] = 2,
                ["Left"] = 3,
                ["Right"] = 4,
                ["Top"] = 5,
                ["Bottom"] = 6
            };

            var headSideOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Left"] = 1,
                ["Right"] = 2
            };

            var poseOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Cheek"] = 1,
                ["Tilt"] = 2
            };

            var channelOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Low"] = 1,
                ["Mid"] = 2,
                ["High"] = 3
            };

            var RbOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["1RB"] = 1,
                ["50%RB"] = 2,
                ["50%"] = 2,
                ["50"] = 2,
                ["100"] = 3,
                ["100%"] = 3,
                ["100%RB"] = 3
            };

            int MatchKey(string name, Dictionary<string, int> map)
            {
                foreach (var k in map.Keys)
                    if (name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                        return map[k];
                return int.MaxValue;
            }

            bool ContainsAny(string text, params string[] patterns)
            {
                foreach (var p in patterns)
                    if (text.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                return false;
            }

            int Get2GFreq(string name)
            {
                if (ContainsAny(name, "GSM 850", "GSM850")) return 850;
                if (ContainsAny(name, "GSM 900", "GSM900")) return 900;
                if (ContainsAny(name, "DCS 1800", "DCS1800", "GSM 1800", "GSM1800")) return 1800;
                if (ContainsAny(name, "PCS 1900", "PCS1900", "GSM 1900", "GSM1900")) return 1900;

                var m = Regex.Match(name, @"(850|900|1800|1900)");
                if (m.Success)
                {
                    switch (m.Value)
                    {
                        case "850": return 850;
                        case "900": return 900;
                        case "1800": return 1800;
                        case "1900": return 1900;
                        default: return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }

            int GetBandNumber(string name)
            {
                var m = Regex.Match(name, @"(?:Band|B)\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            int Get5GBand(string name)
            {
                var m = Regex.Match(name, @"n\s*(\d+)", RegexOptions.IgnoreCase);
                return m.Success ? int.Parse(m.Groups[1].Value) : int.MaxValue;
            }

            int GetRbOrder(string name)
            {
                foreach (var kv in RbOrder)
                {
                    if (name.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        return kv.Value;
                }
                return int.MaxValue;
            }

            // ✅ WLAN / Bluetooth 频率排序
            int GetWlanFreq(string name)
            {
                if (name.IndexOf("2.4G", StringComparison.OrdinalIgnoreCase) >= 0) return 2400;
                if (name.IndexOf("5.2G", StringComparison.OrdinalIgnoreCase) >= 0) return 5200;
                if (name.IndexOf("5.3G", StringComparison.OrdinalIgnoreCase) >= 0) return 5300;
                if (name.IndexOf("5.6G", StringComparison.OrdinalIgnoreCase) >= 0) return 5600;
                if (name.IndexOf("5.8G", StringComparison.OrdinalIgnoreCase) >= 0) return 5800;
                return int.MaxValue;
            }

            return docxFiles
                .Select(f =>
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    var network = MatchKey(name, networkOrder);

                    return new
                    {
                        File = f,
                        Network = network,
                        Region = MatchKey(name, regionOrder),
                        SubPart = MatchKey(name, subPartOrder),
                        HeadSide = MatchKey(name, headSideOrder),
                        Pose = MatchKey(name, poseOrder),
                        Channel = MatchKey(name, channelOrder),
                        Rb = GetRbOrder(name),

                        FreqOrBand = network switch
                        {
                            1 => Get2GFreq(name),
                            2 => GetBandNumber(name),
                            3 => GetBandNumber(name),
                            4 => Get5GBand(name),
                            // ✅ WLAN 按频率，Bluetooth 用一个固定值排在 WLAN 后面
                            5 => GetWlanFreq(name),   // WLAN: 2400 → 5200 → 5300 → 5600 → 5800
                            6 => 6000,                // Bluetooth 排在 WLAN 之后
                            _ => int.MaxValue
                        }
                    };
                })
                .OrderBy(x => x.Network)
                .ThenBy(x => x.FreqOrBand)    // ✅ WLAN 按频率排，Bluetooth 排在最后
                .ThenBy(x => x.Region)
                .ThenBy(x => x.SubPart)
                .ThenBy(x => x.HeadSide)
                .ThenBy(x => x.Pose)
                .ThenBy(x => x.Channel)
                .ThenBy(x => x.Rb)
                .Select(x => x.File)
                .ToArray();
        }
    }


}
