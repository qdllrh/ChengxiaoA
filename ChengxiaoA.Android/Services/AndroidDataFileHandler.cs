using Android.Content;
using AndroidUri = Android.Net.Uri;
using System;
using System.IO;
using System.Text.Json;
using System.Diagnostics;

namespace ChengxiaoA.Android.Services;

/// <summary>
/// Android 平台的数据文件处理器 - 处理 URI 格式的文件路径
/// </summary>
public static class AndroidDataFileHandler
{
    /// <summary>
    /// 保存数据到文件（支持 URI 格式）
    /// </summary>
    public static void SaveDataToFile(object data, string filePathOrUri)
    {
        var context = MainActivity.Current;
        if (context == null)
            throw new InvalidOperationException("无法获取 MainActivity 实例");

        Debug.WriteLine($"[AndroidDataFileHandler] SaveDataToFile 开始, path: {filePathOrUri}");

        try
        {
            // 获取数据属性值
            var dataType = data.GetType();
            var chengxiaoProperty = dataType.GetProperty("Chengxiaoleiji");
            var yuleProperty = dataType.GetProperty("Yuleleiji");
            var duoyuProperty = dataType.GetProperty("Duoyuleiji");

            var chengxiaoValue = chengxiaoProperty?.GetValue(data) ?? 0.0;
            var yuleValue = yuleProperty?.GetValue(data) ?? 0.0;
            var duoyuValue = duoyuProperty?.GetValue(data) ?? 0.0;

            Debug.WriteLine($"[AndroidDataFileHandler] 数据: Chengxiaoleiji={chengxiaoValue}, Yuleleiji={yuleValue}, Duoyuleiji={duoyuValue}");

            // 检查是否是 URI 格式
            if (filePathOrUri.StartsWith("/document/") || filePathOrUri.Contains("content://"))
            {
                // 是 URI 格式，通过 ContentResolver 写入
                AndroidUri uri;
                try
                {
                    // 检查是否是完整 URI
                    if (filePathOrUri.StartsWith("content://"))
                    {
                        // 已经是完整的 content:// URI，直接使用
                        uri = AndroidUri.Parse(filePathOrUri);
                        Debug.WriteLine($"[AndroidDataFileHandler] 使用完整 URI: {uri}");
                    }
                    else if (filePathOrUri.StartsWith("/document/"))
                    {
                        // 是简化的路径，尝试使用文档提供者
                        // 注意：这种格式可能不完整，优先使用 content:// 格式
                        uri = AndroidUri.Parse($"content://com.android.providers.documents.documents{filePathOrUri}");
                        Debug.WriteLine($"[AndroidDataFileHandler] 转换 URI: {filePathOrUri} -> {uri}");
                    }
                    else
                    {
                        throw new ArgumentException($"不支持的 URI 格式: {filePathOrUri}");
                    }
                    
                    Debug.WriteLine($"[AndroidDataFileHandler] 解析 URI 成功: {uri}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] URI 解析失败: {ex.Message}");
                    throw;
                }

                var contentResolver = context.ContentResolver;

                Debug.WriteLine($"[AndroidDataFileHandler] 使用 ContentResolver 写入");

                // 创建 JSON 字符串
                var jsonData = new
                {
                    Chengxiaoleiji = Convert.ToDouble(chengxiaoValue),
                    Yuleleiji = Convert.ToDouble(yuleValue),
                    Duoyuleiji = Convert.ToDouble(duoyuValue)
                };
                var json = JsonSerializer.Serialize(jsonData);
                Debug.WriteLine($"[AndroidDataFileHandler] JSON: {json}");

                // 使用 ContentResolver 写入
                System.IO.Stream? outputStream = null;
                try
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] 尝试打开输出流...");
                    outputStream = contentResolver.OpenOutputStream(uri);
                    
                    if (outputStream != null)
                    {
                        Debug.WriteLine($"[AndroidDataFileHandler] 输出流打开成功");
                        
                        // 转换为字节数组写入（避免编码问题）
                        var bytes = System.Text.Encoding.UTF8.GetBytes(json);
                        outputStream.Write(bytes, 0, bytes.Length);
                        outputStream.Flush();
                        
                        Debug.WriteLine($"[AndroidDataFileHandler] 写入 {bytes.Length} 字节成功");
                    }
                    else
                    {
                        throw new InvalidOperationException($"无法打开输出流: {uri}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] 写入失败: {ex.Message}");
                    Debug.WriteLine($"[AndroidDataFileHandler] 异常类型: {ex.GetType().Name}");
                    Debug.WriteLine($"[AndroidDataFileHandler] 异常堆栈: {ex.StackTrace}");
                    throw;
                }
                finally
                {
                    try
                    {
                        outputStream?.Flush();
                        outputStream?.Close();
                    }
                    catch { }
                    outputStream?.Dispose();
                }
            }
            else
            {
                // 是普通文件路径，直接写入
                var jsonData = new
                {
                    Chengxiaoleiji = Convert.ToDouble(chengxiaoValue),
                    Yuleleiji = Convert.ToDouble(yuleValue),
                    Duoyuleiji = Convert.ToDouble(duoyuValue)
                };
                var json = JsonSerializer.Serialize(jsonData);
                File.WriteAllText(filePathOrUri, json);
                Debug.WriteLine("[AndroidDataFileHandler] 文件写入成功");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[AndroidDataFileHandler] 保存失败: {ex.Message}");
            Debug.WriteLine($"[AndroidDataFileHandler] 异常详情: {ex}");
            throw;
        }
    }

    /// <summary>
    /// 从文件加载数据（支持 URI 格式）
    /// </summary>
    public static object LoadDataFromFile(string filePathOrUri)
    {
        var context = MainActivity.Current;
        if (context == null)
            throw new InvalidOperationException("无法获取 MainActivity 实例");

        Debug.WriteLine($"[AndroidDataFileHandler] LoadDataFromFile 开始, path: {filePathOrUri}");

        try
        {
            string json;

            // 检查是否是 URI 格式
            if (filePathOrUri.StartsWith("/document/") || filePathOrUri.Contains("content://"))
            {
                // 是 URI 格式，通过 ContentResolver 读取
                AndroidUri uri;
                try
                {
                    // 检查是否是完整 URI
                    if (filePathOrUri.StartsWith("content://"))
                    {
                        uri = AndroidUri.Parse(filePathOrUri);
                    }
                    else if (filePathOrUri.StartsWith("/document/"))
                    {
                        // 转换为完整的 content:// URI
                        uri = AndroidUri.Parse($"content://com.android.providers.documents.documents{filePathOrUri}");
                        Debug.WriteLine($"[AndroidDataFileHandler] 转换 URI: {filePathOrUri} -> {uri}");
                    }
                    else
                    {
                        throw new ArgumentException($"不支持的 URI 格式: {filePathOrUri}");
                    }
                    
                    Debug.WriteLine($"[AndroidDataFileHandler] 解析 URI 成功: {uri}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] URI 解析失败: {ex.Message}");
                    throw;
                }

                var contentResolver = context.ContentResolver;

                Debug.WriteLine($"[AndroidDataFileHandler] 使用 ContentResolver 读取");

                System.IO.Stream? inputStream = null;
                try
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] 尝试打开输入流...");
                    inputStream = contentResolver.OpenInputStream(uri);
                    
                    if (inputStream != null)
                    {
                        Debug.WriteLine($"[AndroidDataFileHandler] 输入流打开成功");
                        
                        // 使用字节读取
                        using var memoryStream = new MemoryStream();
                        inputStream.CopyTo(memoryStream);
                        var bytes = memoryStream.ToArray();
                        json = System.Text.Encoding.UTF8.GetString(bytes);
                        
                        Debug.WriteLine($"[AndroidDataFileHandler] 读取 {bytes.Length} 字节");
                        Debug.WriteLine($"[AndroidDataFileHandler] 读取到的 JSON: {json}");
                    }
                    else
                    {
                        throw new FileNotFoundException($"无法打开文件: {filePathOrUri}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AndroidDataFileHandler] 读取失败: {ex.Message}");
                    Debug.WriteLine($"[AndroidDataFileHandler] 异常类型: {ex.GetType().Name}");
                    Debug.WriteLine($"[AndroidDataFileHandler] 异常堆栈: {ex.StackTrace}");
                    throw;
                }
                finally
                {
                    try
                    {
                        inputStream?.Close();
                    }
                    catch { }
                    inputStream?.Dispose();
                }
            }
            else
            {
                // 是普通文件路径，直接读取
                json = File.ReadAllText(filePathOrUri);
                Debug.WriteLine($"[AndroidDataFileHandler] 读取到的 JSON: {json}");
            }

            // 解析 JSON
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var chengxiaoValue = root.TryGetProperty("Chengxiaoleiji", out var chengxiaoElement) ? chengxiaoElement.GetDouble() : 0.0;
            var yuleValue = root.TryGetProperty("Yuleleiji", out var yuleElement) ? yuleElement.GetDouble() : 0.0;
            var duoyuValue = root.TryGetProperty("Duoyuleiji", out var duoyuElement) ? duoyuElement.GetDouble() : 0.0;

            Debug.WriteLine($"[AndroidDataFileHandler] 解析数据: Chengxiaoleiji={chengxiaoValue}, Yuleleiji={yuleValue}, Duoyuleiji={duoyuValue}");

            // 创建匿名对象返回
            return new { Chengxiaoleiji = chengxiaoValue, Yuleleiji = yuleValue, Duoyuleiji = duoyuValue };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[AndroidDataFileHandler] 加载失败: {ex.Message}");
            Debug.WriteLine($"[AndroidDataFileHandler] 异常详情: {ex}");
            throw;
        }
    }

    /// <summary>
    /// 判断是否是 URI 格式
    /// </summary>
    public static bool IsUriFormat(string path)
    {
        return path.StartsWith("/document/") || path.Contains(":") && !path.Contains("\\") && !path.Contains("/");
    }
}
