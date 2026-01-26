using Avalonia.Platform.Storage;
using ChengxiaoA.Models;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace ChengxiaoA.Services;

/// <summary>
/// 文件处理类 - 支持跨平台文件保存和加载
/// </summary>
public static class DataFileHandler
{
    public static void SaveDataToFile<T>(T data, string filePath)
    {
        // 检测是否是 URI 格式（Android 返回的路径格式）
        // content:// 开头的 URI（完整 URI）
        // /document/ 开头的简化 URI
        bool isUriFormat = filePath.StartsWith("content://") || filePath.StartsWith("/document/");

        if (isUriFormat)
        {
            // 是 URI 格式，使用 JSON 格式，并通过反射调用 Android 专用的处理器
            Debug.WriteLine($"检测到 URI 格式，准备保存: {filePath}");
            
            // 调用 Android 专用的处理器
            try
            {
                // 列出所有程序集，查找类型
                Debug.WriteLine("开始查找 AndroidDataFileHandler 类型...");
                
                var handlerType = System.Type.GetType("ChengxiaoA.Android.Services.AndroidDataFileHandler, ChengxiaoA.Android");
                
                if (handlerType == null)
                {
                    Debug.WriteLine("尝试使用简单类型名查找...");
                    handlerType = System.Type.GetType("ChengxiaoA.Android.Services.AndroidDataFileHandler");
                }
                
                if (handlerType == null)
                {
                    Debug.WriteLine("尝试从所有程序集中查找...");
                    foreach (var assembly in System.AppDomain.CurrentDomain.GetAssemblies())
                    {
                        Debug.WriteLine($"检查程序集: {assembly.GetName().Name}");
                        if (assembly.GetName().Name?.Contains("ChengxiaoA.Android") == true)
                        {
                            Debug.WriteLine($"  找到匹配的程序集");
                            handlerType = assembly.GetType("ChengxiaoA.Android.Services.AndroidDataFileHandler");
                            if (handlerType != null)
                            {
                                Debug.WriteLine($"  找到类型!");
                                break;
                            }
                        }
                    }
                }
                
                if (handlerType != null)
                {
                    Debug.WriteLine($"成功找到 AndroidDataFileHandler 类型: {handlerType.FullName}");
                    Debug.WriteLine($"类型程序集: {handlerType.Assembly.GetName().Name}");
                    
                    var saveMethod = handlerType.GetMethod("SaveDataToFile", BindingFlags.Public | BindingFlags.Static);
                    if (saveMethod != null)
                    {
                        Debug.WriteLine($"找到 SaveDataToFile 方法，参数: {string.Join(", ", saveMethod.GetParameters().Select(p => p.ParameterType.Name))}");
                        Debug.WriteLine("调用 AndroidDataFileHandler.SaveDataToFile");
                        saveMethod.Invoke(null, new object[] { data, filePath });
                        Debug.WriteLine("保存完成");
                    }
                    else
                    {
                        Debug.WriteLine("未找到 SaveDataToFile 方法");
                        // 列出所有方法
                        Debug.WriteLine("可用方法:");
                        foreach (var method in handlerType.GetMethods(BindingFlags.Public | BindingFlags.Static))
                        {
                            Debug.WriteLine($"  - {method.Name}");
                        }
                    }
                }
                else
                {
                    Debug.WriteLine("未找到 AndroidDataFileHandler 类型");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Android 保存失败: {ex.Message}");
                Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                if (ex.InnerException != null)
                {
                    Debug.WriteLine($"内部异常: {ex.InnerException.Message}");
                    Debug.WriteLine($"内部异常类型: {ex.InnerException.GetType().Name}");
                    Debug.WriteLine($"内部异常堆栈: {ex.InnerException.StackTrace}");
                }
                Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
            }
        }
        else
        {
            // 其他平台：使用 JSON 格式
            Debug.WriteLine($"使用普通文件路径: {filePath}");
            var json = JsonSerializer.Serialize(data);
            File.WriteAllText(filePath, json);
            Debug.WriteLine("文件写入成功");
        }
    }

    public static T LoadDataFromFile<T>(string filePath) where T : new()
    {
        var result = new T();

        // 检测是否是 URI 格式（Android 返回的路径格式）
        // content:// 开头的 URI（完整 URI）
        // /document/ 开头的简化 URI
        bool isUriFormat = filePath.StartsWith("content://") || filePath.StartsWith("/document/");

        if (isUriFormat)
        {
            // 是 URI 格式，使用 JSON 格式，并通过反射调用 Android 专用的处理器
            Debug.WriteLine($"检测到 URI 格式，准备加载: {filePath}");
            
            try
            {
                var handlerType = System.Type.GetType("ChengxiaoA.Android.Services.AndroidDataFileHandler, ChengxiaoA.Android");
                if (handlerType != null)
                {
                    Debug.WriteLine("成功找到 AndroidDataFileHandler 类型");
                    var loadMethod = handlerType.GetMethod("LoadDataFromFile", BindingFlags.Public | BindingFlags.Static);
                    if (loadMethod != null)
                    {
                        Debug.WriteLine("调用 AndroidDataFileHandler.LoadDataFromFile");
                        var loadedData = loadMethod.Invoke(null, new object[] { filePath });

                        // 使用 JSON 反序列化
                        if (loadedData != null)
                        {
                            // 获取数据类型
                            var jsonOptions = new JsonSerializerOptions
                            {
                                PropertyNameCaseInsensitive = true
                            };

                            // 将加载的数据转换为 JSON 字符串，然后反序列化到目标类型
                            var json = JsonSerializer.Serialize(loadedData);
                            result = JsonSerializer.Deserialize<T>(json, jsonOptions) ?? new T();
                        }
                    }
                    else
                    {
                        Debug.WriteLine("未找到 LoadDataFromFile 方法");
                    }
                }
                else
                {
                    Debug.WriteLine("未找到 AndroidDataFileHandler 类型");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Android 加载失败: {ex.Message}");
                Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                if (ex.InnerException != null)
                {
                    Debug.WriteLine($"内部异常: {ex.InnerException.Message}");
                    Debug.WriteLine($"内部异常类型: {ex.InnerException.GetType().Name}");
                    Debug.WriteLine($"内部异常堆栈: {ex.InnerException.StackTrace}");
                }
                Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
            }
        }
        else
        {
            // 其他平台：使用 JSON 格式
            Debug.WriteLine($"使用普通文件路径: {filePath}");
            var json = File.ReadAllText(filePath);
            result = JsonSerializer.Deserialize<T>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new T();
        }

        return result;
    }
}
