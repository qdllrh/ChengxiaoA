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
            if (filePathOrUri.StartsWith("/document/") || filePathOrUri.Contains(":"))
            {
                // 是 URI 格式，通过 ContentResolver 写入
                AndroidUri uri;
                try
                {
                    uri = AndroidUri.Parse(filePathOrUri);
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
                    outputStream = contentResolver.OpenOutputStream(uri);
                    if (outputStream != null)
                    {
                        using var writer = new StreamWriter(outputStream);
                        writer.Write(json);
                        writer.Flush();
                        Debug.WriteLine("[AndroidDataFileHandler] 写入成功");
                    }
                    else
                    {
                        throw new InvalidOperationException($"无法打开输出流: {uri}");
                    }
                }
                finally
                {
                    outputStream?.Close();
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
            if (filePathOrUri.StartsWith("/document/") || filePathOrUri.Contains(":"))
            {
                // 是 URI 格式，通过 ContentResolver 读取
                AndroidUri uri;
                try
                {
                    uri = AndroidUri.Parse(filePathOrUri);
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
                    inputStream = contentResolver.OpenInputStream(uri);
                    if (inputStream != null)
                    {
                        using var reader = new StreamReader(inputStream);
                        json = reader.ReadToEnd();
                        Debug.WriteLine($"[AndroidDataFileHandler] 读取到的 JSON: {json}");
                    }
                    else
                    {
                        throw new FileNotFoundException($"无法打开文件: {filePathOrUri}");
                    }
                }
                finally
                {
                    inputStream?.Close();
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
