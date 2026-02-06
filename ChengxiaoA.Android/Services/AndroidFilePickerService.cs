using Android.App;
using Android.Content;
using Android.OS;
using Android.Provider;
using ChengxiaoA.Services;
using System;
using System.Threading.Tasks;

namespace ChengxiaoA.Android.Services;

/// <summary>
/// Android 系统文件选择器服务（获取带权限的文件 URI）
/// </summary>
public class AndroidFilePickerService : IFilePickerService
{
    // 静态变量存储回调和请求码
    private static TaskCompletionSource<string?> _filePickerTcs;
    private const int FilePickerRequestCode = 1001;

    /// <summary>
    /// 打开系统文件选择器，选择 Excel 文件
    /// </summary>
    public async Task<string> PickExcelFileAsync()
    {
        _filePickerTcs = new TaskCompletionSource<string?>();

        var activity = MainActivity.Current;
        if (activity == null)
        {
            System.Diagnostics.Debug.WriteLine("MainActivity 为空，无法打开文件选择器");
            return string.Empty;
        }

        try
        {
            // 1. 创建文件选择意图（ACTION_OPEN_DOCUMENT 是关键，能获取权限）
            var intent = new Intent(Intent.ActionOpenDocument);
            intent.AddCategory(Intent.CategoryOpenable); // 只允许可打开的文件
            intent.SetType("application/vnd.ms-excel"); // .xls
            intent.PutExtra(Intent.ExtraMimeTypes, new[] {
                "application/vnd.ms-excel",       // .xls
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" // .xlsx
            });

            // 2. 启动文件选择器（等待用户选择）
            activity.StartActivityForResult(intent, FilePickerRequestCode);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"打开文件选择器失败：{ex.Message}");
            _filePickerTcs.TrySetResult(null);
        }

        // 3. 等待用户选择文件，返回带权限的 URI
        var result = await _filePickerTcs.Task;
        return result ?? string.Empty;
    }

    /// <summary>
    /// 处理文件选择器的返回结果（需要在 MainActivity 中调用）
    /// </summary>
    public static void OnActivityResult(int requestCode, Result resultCode, Intent? data)
    {
        if (requestCode != FilePickerRequestCode || _filePickerTcs == null)
            return;

        try
        {
            if (resultCode == Result.Ok && data != null && data.Data != null)
            {
                var uri = data.Data.ToString();
                System.Diagnostics.Debug.WriteLine($"文件选择成功，URI: {uri}");

                // 关键：获取 URI 的持久化权限（避免后续访问失效）
                var activity = MainActivity.Current;
                if (activity != null)
                {
                    var takeFlags = data.Flags & (ActivityFlags.GrantReadUriPermission | ActivityFlags.GrantWriteUriPermission);
                    activity.ContentResolver.TakePersistableUriPermission(
                        data.Data, takeFlags);
                }

                _filePickerTcs.TrySetResult(uri);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("用户取消选择或选择失败");
                _filePickerTcs.TrySetResult(null);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"处理文件选择结果失败：{ex.Message}");
            _filePickerTcs.TrySetResult(null);
        }
    }

    /// <summary>
    /// 保存文件（Android 暂未实现）
    /// </summary>
    public Task<string> SaveFileAsync(string defaultFileName)
    {
        // TODO: 实现 Android 保存文件对话框
        return Task.FromResult(string.Empty);
    }
}
