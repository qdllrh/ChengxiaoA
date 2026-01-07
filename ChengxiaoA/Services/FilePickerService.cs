using System;
using System.Threading.Tasks;

namespace ChengxiaoA.Services;

/// <summary>
/// 文件选择器服务 - 默认实现（用于非 Android 平台）
/// </summary>
public class FilePickerService
{
    private static IFilePickerService? _instance;

    /// <summary>
    /// 获取文件选择器实例（平台特定）
    /// </summary>
    public static IFilePickerService Instance
    {
        get
        {
            _instance ??= new DefaultFilePickerService();
            return _instance;
        }
        set
        {
            _instance = value;
        }
    }

    /// <summary>
    /// 选择 Excel 文件（便捷方法）
    /// </summary>
    public static Task<string> PickExcelFileAsync()
    {
        return Instance.PickExcelFileAsync();
    }

    /// <summary>
    /// 保存文件（便捷方法）
    /// </summary>
    public static Task<string> SaveFileAsync(string defaultFileName)
    {
        return Instance.SaveFileAsync(defaultFileName);
    }

    /// <summary>
    /// 默认文件选择器实现（非 Android 平台）
    /// </summary>
    private class DefaultFilePickerService : IFilePickerService
    {
        public Task<string> PickExcelFileAsync()
        {
            // 非平台特定的实现，返回空字符串
            // 实际使用时应该由具体平台实现
            return Task.FromResult(string.Empty);
        }

        public Task<string> SaveFileAsync(string defaultFileName)
        {
            // 非平台特定的实现，返回空字符串
            return Task.FromResult(string.Empty);
        }
    }
}
