// Android 子项目 - ChengxiaoA.Android/Services/AndroidExcelReaderService.cs
using System;
using System.IO;
using System.Net;
using Android.Content;
using Android.Net;
using ChengxiaoA.Services; // 引用主项目的接口
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ChengxiaoA.Android.Services
{
    /// <summary>
    /// Android 平台的 Excel 读取服务（实现主项目接口）
    /// </summary>
    public class AndroidExcelReaderService : IExcelReaderService
    {
        // 实现接口的核心方法（就是你改造后的 ReadFirstCellNumber）
        public double ReadFirstCellNumber(string filePath)
        {
            IWorkbook workbook = null;
            Stream? dataStream = null;

            try
            {
                System.Diagnostics.Debug.WriteLine($"=== ReadFirstCellNumber 开始 ===");
                System.Diagnostics.Debug.WriteLine($"文件路径: {filePath}");

                // 检查是否是 Android URI（content:// 开头或 /document/ 开头）
                bool isAndroidUri = filePath.StartsWith("content://") ||
                                   filePath.StartsWith("/document/");

                if (isAndroidUri)
                {
                    System.Diagnostics.Debug.WriteLine("检测到 Android URI 格式，调用 GetAndroidUriStream");
                    dataStream = GetAndroidUriStream(filePath);
                    if (dataStream == null)
                    {
                        System.Diagnostics.Debug.WriteLine("GetAndroidUriStream 返回 null");
                        return 0;
                    }
                    System.Diagnostics.Debug.WriteLine("GetAndroidUriStream 返回成功");
                }

                else
                {
                    if (!File.Exists(filePath)) return 0;
                    dataStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                }

                string extension = GetFileExtension(filePath).ToLower();
                if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook(dataStream);
                }
                else if (extension == ".xlsx")
                {
                    workbook = new XSSFWorkbook(dataStream);
                }
                else
                {
                    return 0;
                }

                ISheet sheet = workbook.GetSheetAt(0);
                IRow row = sheet.GetRow(0);
                if (row == null) return 0;

                ICell cell = row.GetCell(0);
                if (cell == null) return 0;

                double value = 0;
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        value = cell.NumericCellValue;
                        break;
                    case CellType.String:
                        if (double.TryParse(cell.StringCellValue, out double num))
                            value = num;
                        else
                            throw new Exception($"第一个单元格不是有效的数字: {cell.StringCellValue}");
                        break;
                    case CellType.Formula:
                        if (cell.CachedFormulaResultType == CellType.Numeric)
                            value = cell.NumericCellValue;
                        else if (cell.CachedFormulaResultType == CellType.String)
                        {
                            if (double.TryParse(cell.StringCellValue, out double num2))
                                value = num2;
                            else
                                throw new Exception($"第一个单元格公式结果不是有效的数字: {cell.StringCellValue}");
                        }
                        else
                            throw new Exception($"第一个单元格公式结果类型不支持: {cell.CachedFormulaResultType}");
                        break;
                    case CellType.Blank:
                        return 0;
                    default:
                        throw new Exception($"第一个单元格类型不支持: {cell.CellType}");
                }

                return value;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"读取Excel出错：{ex.Message}");
                return 0;
            }
            finally
            {
                workbook?.Close();
                workbook?.Dispose();
                dataStream?.Close();
                dataStream?.Dispose();
            }
        }

        public bool WriteFirstCellNumber(string filePath, double value)
        {
            IWorkbook workbook = null;
            Stream? dataStream = null;
            Stream? outputStream = null;

            try
            {
                System.Diagnostics.Debug.WriteLine($"=== WriteFirstCellNumber 开始 ===");
                System.Diagnostics.Debug.WriteLine($"文件路径: {filePath}");
                System.Diagnostics.Debug.WriteLine($"写入数值: {value}");

                // 检查是否是 Android URI
                bool isAndroidUri = filePath.StartsWith("content://") ||
                                   filePath.StartsWith("/document/");

                // 1. 读取现有文件（如果是新文件则需要创建）
                string extension = GetFileExtension(filePath).ToLower();

                if (isAndroidUri)
                {
                    // Android URI 先尝试读取现有文件
                    dataStream = GetAndroidUriStream(filePath);
                    if (dataStream == null)
                    {
                        System.Diagnostics.Debug.WriteLine("文件不存在，创建新工作簿");
                        // 创建新工作簿
                        if (extension == ".xls")
                            workbook = new HSSFWorkbook();
                        else if (extension == ".xlsx")
                            workbook = new XSSFWorkbook();
                        else
                            return false;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("读取现有文件");
                        // 加载现有工作簿
                        if (extension == ".xls")
                            workbook = new HSSFWorkbook(dataStream);
                        else if (extension == ".xlsx")
                            workbook = new XSSFWorkbook(dataStream);
                        else
                            return false;
                    }
                }
                else
                {
                    // 普通文件路径
                    if (File.Exists(filePath))
                    {
                        dataStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                        if (extension == ".xls")
                            workbook = new HSSFWorkbook(dataStream);
                        else if (extension == ".xlsx")
                            workbook = new XSSFWorkbook(dataStream);
                        else
                            return false;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("文件不存在，创建新工作簿");
                        if (extension == ".xls")
                            workbook = new HSSFWorkbook();
                        else if (extension == ".xlsx")
                            workbook = new XSSFWorkbook();
                        else
                            return false;
                    }
                }

                // 2. 写入第一个单元格
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }

                IRow row = sheet.GetRow(0);
                if (row == null)
                {
                    row = sheet.CreateRow(0);
                }

                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                }

                cell.SetCellValue(value);
                System.Diagnostics.Debug.WriteLine($"已写入值 {value} 到第一个单元格");

                // 3. 保存文件
                MemoryStream memoryStream = new MemoryStream();
                workbook.Write(memoryStream);
                memoryStream.Position = 0;

                if (isAndroidUri)
                {
                    // Android URI 写入
                    System.Diagnostics.Debug.WriteLine("通过 Android URI 写入文件");
                    outputStream = GetAndroidUriWriteStream(filePath);
                    if (outputStream == null)
                    {
                        System.Diagnostics.Debug.WriteLine("无法获取写入流");
                        return false;
                    }
                    memoryStream.CopyTo(outputStream);
                    outputStream.Flush();
                    System.Diagnostics.Debug.WriteLine("文件写入成功");
                }
                else
                {
                    // 普通文件路径写入
                    System.Diagnostics.Debug.WriteLine($"通过文件路径写入: {filePath}");
                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        memoryStream.CopyTo(fileStream);
                    }
                    System.Diagnostics.Debug.WriteLine("文件写入成功");
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"写入Excel出错：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                return false;
            }
            finally
            {
                workbook?.Close();
                workbook?.Dispose();
                dataStream?.Close();
                dataStream?.Dispose();
                outputStream?.Close();
                outputStream?.Dispose();
            }
        }

        // 辅助方法1：获取 Android URI 数据流（复用你的代码）
        private Stream? GetAndroidUriWriteStream(string uriPath)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== GetAndroidUriWriteStream 开始 ===");
                System.Diagnostics.Debug.WriteLine($"原始 URI 路径: {uriPath}");

                // 1. URL 解码
                string decodedUri = WebUtility.UrlDecode(uriPath);
                System.Diagnostics.Debug.WriteLine($"URL 解码后的 URI: {decodedUri}");

                global::Android.Net.Uri? uri = null;

                // 2. 检测是否是嵌套 raw: 格式的 URI
                if (decodedUri.Contains("/document/raw:/"))
                {
                    System.Diagnostics.Debug.WriteLine("检测到嵌套 raw: 格式，转换为标准 URI");

                    int rawIndex = decodedUri.IndexOf("/document/raw:/");
                    string rawPathPart = decodedUri.Substring(rawIndex + "/document/".Length);
                    System.Diagnostics.Debug.WriteLine($"Raw 路径部分: {rawPathPart}");

                    string standardUri = $"content://com.android.providers.downloads.documents/document/{WebUtility.UrlEncode(rawPathPart)}";
                    uri = global::Android.Net.Uri.Parse(standardUri);
                    System.Diagnostics.Debug.WriteLine($"转换后的标准 URI: {standardUri}");
                }
                else if (uriPath.StartsWith("/document/"))
                {
                    string decodedPath = WebUtility.UrlDecode(uriPath);
                    string docId = decodedPath.Replace("/document/", "");
                    decodedPath = $"content://com.android.providers.downloads.documents/document/{docId}";
                    uri = global::Android.Net.Uri.Parse(decodedPath);
                    System.Diagnostics.Debug.WriteLine($"补全后的 URI: {decodedPath}");
                }
                else
                {
                    uri = global::Android.Net.Uri.Parse(uriPath);
                    System.Diagnostics.Debug.WriteLine($"使用原始 URI");
                }

                if (uri == null)
                {
                    System.Diagnostics.Debug.WriteLine("URI 解析失败");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"URI 解析成功: {uri}");

                if (string.IsNullOrEmpty(uri.Scheme) || string.IsNullOrEmpty(uri.Host))
                {
                    System.Diagnostics.Debug.WriteLine($"解析后的 URI 无效（缺少 Scheme/Host）");
                    return null;
                }

                Context? context = MainActivity.Current;
                if (context == null)
                {
                    System.Diagnostics.Debug.WriteLine("Android 上下文为空");
                    return null;
                }

                // 3. 持久化 URI 权限（写入也需要权限）
                try
                {
                    var takeFlags = global::Android.Content.ActivityFlags.GrantReadUriPermission | global::Android.Content.ActivityFlags.GrantWriteUriPermission;
                    context.ContentResolver.TakePersistableUriPermission(uri, takeFlags);
                    System.Diagnostics.Debug.WriteLine("✅ URI 持久化权限激活成功（读写）");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ 持久化权限失败（临时权限可用）: {ex.Message}");
                }

                // 4. 打开输出流
                var outputStream = context.ContentResolver.OpenOutputStream(uri);
                if (outputStream == null)
                {
                    System.Diagnostics.Debug.WriteLine("无法打开 URI 输出流");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine("输出流打开成功");
                return outputStream;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取 URI 写入流出错：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                System.Diagnostics.Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
                return null;
            }
        }

        private Stream? GetAndroidUriStream(string uriPath)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== GetAndroidUriStream 开始 ===");
                System.Diagnostics.Debug.WriteLine($"原始 URI 路径: {uriPath}");

                // 1. URL 解码
                string decodedUri = WebUtility.UrlDecode(uriPath);
                System.Diagnostics.Debug.WriteLine($"URL 解码后的 URI: {decodedUri}");

                global::Android.Net.Uri? uri = null;

                // 2. 检测是否是嵌套 raw: 格式的 URI
                // 格式: content://com.android.providers.downloads.documents/document/raw:/storage/...
                // 这种格式需要将 raw:/storage/... 转换为标准的 document ID 格式
                if (decodedUri.Contains("/document/raw:/"))
                {
                    System.Diagnostics.Debug.WriteLine("检测到嵌套 raw: 格式，转换为标准 URI");

                    // 提取 raw: 后面的路径部分
                    int rawIndex = decodedUri.IndexOf("/document/raw:/");
                    string rawPathPart = decodedUri.Substring(rawIndex + "/document/".Length);
                    System.Diagnostics.Debug.WriteLine($"Raw 路径部分: {rawPathPart}");

                    // 构建标准的 content URI，使用 DownloadsProvider 的 raw: 格式支持
                    // 对于 DownloadsProvider，文档 ID 可以是 raw:/storage/... 的格式
                    string standardUri = $"content://com.android.providers.downloads.documents/document/{WebUtility.UrlEncode(rawPathPart)}";
                    uri = global::Android.Net.Uri.Parse(standardUri);
                    System.Diagnostics.Debug.WriteLine($"转换后的标准 URI: {standardUri}");
                }
                else if (uriPath.StartsWith("/document/"))
                {
                    // 3. 旧格式：/document/ 开头，补全为 content:// URI
                    string decodedPath = WebUtility.UrlDecode(uriPath);
                    string docId = decodedPath.Replace("/document/", "");
                    decodedPath = $"content://com.android.providers.downloads.documents/document/{docId}";
                    uri = global::Android.Net.Uri.Parse(decodedPath);
                    System.Diagnostics.Debug.WriteLine($"补全后的 URI: {decodedPath}");
                }
                else
                {
                    // 4. 直接是 content:// URI，直接解析
                    uri = global::Android.Net.Uri.Parse(uriPath);
                    System.Diagnostics.Debug.WriteLine($"使用原始 URI");
                }

                if (uri == null)
                {
                    System.Diagnostics.Debug.WriteLine("URI 解析失败");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"URI 解析成功: {uri}");
                System.Diagnostics.Debug.WriteLine($"URI Scheme: {uri.Scheme}");
                System.Diagnostics.Debug.WriteLine($"URI Host: {uri.Host}");

                // 5. 校验 URI 是否有效
                if (string.IsNullOrEmpty(uri.Scheme) || string.IsNullOrEmpty(uri.Host))
                {
                    System.Diagnostics.Debug.WriteLine($"解析后的 URI 无效（缺少 Scheme/Host）");
                    return null;
                }

                // 6. 通过 ContentResolver 获取输入流
                Context? context = MainActivity.Current;
                System.Diagnostics.Debug.WriteLine($"MainActivity.Current: {context?.GetType().Name ?? "null"}");
                if (context == null)
                {
                    System.Diagnostics.Debug.WriteLine("Android 上下文为空");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"ContentResolver: {context.ContentResolver?.GetType().Name ?? "null"}");

                // 7. 持久化 URI 权限（关键！）
                try
                {
                    var takeFlags = global::Android.Content.ActivityFlags.GrantReadUriPermission;
                    context.ContentResolver.TakePersistableUriPermission(uri, takeFlags);
                    System.Diagnostics.Debug.WriteLine("✅ URI 持久化权限激活成功");
                }
                catch (Exception ex)
                {
                    // 非致命异常：如果是临时权限，跳过即可（系统已自动授予临时权限）
                    System.Diagnostics.Debug.WriteLine($"⚠️ 持久化权限失败（临时权限可用）: {ex.Message}");
                }

                // 8. 打开输入流
                var inputStream = context.ContentResolver.OpenInputStream(uri);
                if (inputStream == null)
                {
                    System.Diagnostics.Debug.WriteLine("无法打开 URI 输入流");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine("输入流打开成功");

                // 9. 转换为 MemoryStream（NPOI 兼容）
                MemoryStream memoryStream = new MemoryStream();
                inputStream.CopyTo(memoryStream);
                memoryStream.Position = 0; // 重置流指针到开头
                inputStream.Close();

                System.Diagnostics.Debug.WriteLine($"流转换成功，大小: {memoryStream.Length} 字节");
                return memoryStream;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取 URI 数据流出错：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                System.Diagnostics.Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
                return null;
            }
        }

        // 辅助方法2：解析文件扩展名（复用你的代码）
        //private Stream? GetAndroidUriStream(string uriPath)
        //{
        //    try
        //    {
        //        System.Diagnostics.Debug.WriteLine($"=== GetAndroidUriStream 开始 ===\n原始 URI 路径: {uriPath}");

        //        // ========== 全新升级：兼容 编码/未编码、简化/完整 所有raw格式URI ==========
        //        // 步骤1：先对原始URI做URL解码（关键！解决%3A/%2F编码问题）
        //        string decodedUri = System.Net.WebUtility.UrlDecode(uriPath);
        //        System.Diagnostics.Debug.WriteLine($"URL解码后的URI: {decodedUri}");

        //        // 步骤2：判断解码后的URI是否包含raw:/storage/（此时编码的%3A/%2F已经还原成:/）
        //        if (decodedUri.Contains("raw:/storage/"))
        //        {
        //            System.Diagnostics.Debug.WriteLine("检测到 raw:+本地路径 格式的URI，开始转换为本地路径");
        //            // 步骤3：提取raw:后面的真实本地路径（解码后直接截取，无需处理编码）
        //            int rawIndex = decodedUri.IndexOf("raw:/"); // 找到raw:/的起始位置
        //            string localPath = decodedUri.Substring(rawIndex); // 截取：raw:/storage/.../试验1.xls
        //            localPath = localPath.Replace("raw:", ""); // 去掉raw:，得到纯本地路径

        //            // 直接用本地路径打开FileStream（和上位机逻辑一致，绕开所有Android坑）
        //            if (System.IO.File.Exists(localPath))
        //            {
        //                var fileStream = new System.IO.FileStream(localPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //                System.Diagnostics.Debug.WriteLine($"✅ 转换成功！本地路径：{localPath}\n✅ FileStream打开成功，流可用");
        //                return fileStream;
        //            }
        //            else
        //            {
        //                System.Diagnostics.Debug.WriteLine($"❌ 转换后的本地路径不存在：{localPath}");
        //                return null;
        //            }
        //        }
        //        // ========== 原有URI解析逻辑（不变，兼容标准Content URI） ==========
        //        global::Android.Net.Uri? uri = null;
        //        if (uriPath.StartsWith("/document/"))
        //        {
        //            string decodedPath = System.Net.WebUtility.UrlDecode(uriPath);
        //            string docId = decodedPath.Replace("/document/", "");
        //            decodedPath = $"content://com.android.providers.downloads.documents/document/{docId}";
        //            uri = global::Android.Net.Uri.Parse(decodedPath);
        //            System.Diagnostics.Debug.WriteLine($"补全后的标准Content URI: {decodedPath}");
        //        }
        //        else
        //        {
        //            uri = global::Android.Net.Uri.Parse(uriPath);
        //            System.Diagnostics.Debug.WriteLine($"使用原始标准Content URI");
        //        }

        //        if (uri == null || string.IsNullOrEmpty(uri.Scheme) || string.IsNullOrEmpty(uri.Host))
        //        {
        //            System.Diagnostics.Debug.WriteLine($"❌ URI 解析无效（缺少Scheme/Host）");
        //            return null;
        //        }
        //        System.Diagnostics.Debug.WriteLine($"✅ URI 解析成功: {uri}");

        //        global::Android.Content.Context? context = ChengxiaoA.Android.MainActivity.Current;
        //        if (context == null)
        //        {
        //            System.Diagnostics.Debug.WriteLine("❌ Android 上下文为空（MainActivity.Current=null）");
        //            return null;
        //        }

        //        // 跳过所有权限代码（关键，避免再次触发异常）
        //        System.Diagnostics.Debug.WriteLine("⚠️ 跳过权限激活，使用系统临时权限");

        //        // 打开输入流（适配标准Content URI）
        //        var inputStream = context.ContentResolver.OpenInputStream(uri);
        //        if (inputStream == null)
        //        {
        //            System.Diagnostics.Debug.WriteLine("❌ 无法打开标准URI的输入流");
        //            return null;
        //        }
        //        System.Diagnostics.Debug.WriteLine("✅ 标准URI输入流打开成功");

        //        // 转换为MemoryStream（NPOI兼容）
        //        System.IO.MemoryStream memoryStream = new System.IO.MemoryStream();
        //        inputStream.CopyTo(memoryStream);
        //        memoryStream.Position = 0; // 重置流指针到开头
        //        inputStream.Close();

        //        System.Diagnostics.Debug.WriteLine($"✅ 流转换成功，MemoryStream大小: {memoryStream.Length} 字节");
        //        return memoryStream;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine($"❌ 获取URI数据流出错：{ex.Message}");
        //        System.Diagnostics.Debug.WriteLine($"❌ 异常类型: {ex.GetType().Name}");
        //        System.Diagnostics.Debug.WriteLine($"❌ 异常堆栈: {ex.StackTrace}");
        //        return null;
        //    }
        //}
        private string GetFileExtension(string filePath)
        {
            // 1. 先解码 URI
            string decodedPath = WebUtility.UrlDecode(filePath);

            // ========== 核心修复：替换 StringSplitOptions.Last 为兼容写法 ==========
            // 原错误写法：decodedPath.Split(new[] { "/document/" }, StringSplitOptions.Last)[1];
            // 兼容写法：分割后取最后一个非空元素
            if (decodedPath.Contains("/document/"))
            {
                string[] parts = decodedPath.Split(new[] { "/document/" }, StringSplitOptions.RemoveEmptyEntries);
                // 取分割后的最后一部分（达到和 Last 一样的效果）
                decodedPath = parts.Length > 0 ? parts[parts.Length - 1] : decodedPath;
            }
            if (decodedPath.Contains("raw:"))
            {
                string[] parts = decodedPath.Split(new[] { "raw:" }, StringSplitOptions.RemoveEmptyEntries);
                decodedPath = parts.Length > 0 ? parts[parts.Length - 1] : decodedPath;
            }

            // 3. 获取扩展名
            return Path.GetExtension(decodedPath);
        }
        public static class AndroidServiceContainer
        {
            // 静态存储 Excel 读取服务实例
            public static IExcelReaderService? ExcelReaderService { get; set; }
        }
    }
}