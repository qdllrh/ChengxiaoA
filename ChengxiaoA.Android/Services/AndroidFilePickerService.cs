using Android.App;
using Android.Content;
using Android.OS;
using Android.Provider;
using AndroidX.Activity.Result;
using AndroidX.Activity.Result.Contract;
using AndroidX.Fragment.App;
using Avalonia.Android;
using ChengxiaoA.Services;
using AndroidUri = Android.Net.Uri;
using System;
using System.Threading.Tasks;

namespace ChengxiaoA.Android.Services;

/// <summary>
/// Android 平台的文件选择器实现
/// </summary>
public class AndroidFilePickerService : IFilePickerService
{
    private static TaskCompletionSource<string>? _tcs;
    private static readonly ActivityResultLauncher _pickerLauncher;

    static AndroidFilePickerService()
    {
        try
        {
            // 获取 MainActivity 实例
            var activity = MainActivity.Current;
            if (activity == null)
            {
                System.Diagnostics.Debug.WriteLine("无法获取 MainActivity 实例");
                return;
            }

            // 注册文件选择器
            _pickerLauncher = activity.RegisterForActivityResult(
                new ActivityResultContracts.GetContent(),
                new ActivityResultCallback()
            );
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"初始化文件选择器失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 选择 Excel 文件
    /// </summary>
    public async Task<string> PickExcelFileAsync()
    {
        _tcs = new TaskCompletionSource<string>();

        try
        {
            var activity = MainActivity.Current;
            if (activity == null)
            {
                System.Diagnostics.Debug.WriteLine("无法获取 MainActivity 实例");
                return string.Empty;
            }

            // 创建选择 Excel 文件的 Intent
            var intent = new Intent(Intent.ActionOpenDocument);
            intent.AddCategory(Intent.CategoryOpenable);
            intent.SetType("*/*");

            // 添加 MIME 类型过滤
            var mimeTypes = new string[]
            {
                "application/vnd.ms-excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/octet-stream"
            };
            intent.PutExtra(Intent.ExtraMimeTypes, mimeTypes);

            // 启动文件选择器
            _pickerLauncher.Launch(intent);

            // 等待用户选择
            return await _tcs.Task;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"文件选择失败: {ex.Message}");
            return string.Empty;
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

    /// <summary>
    /// 结果回调
    /// </summary>
    private class ActivityResultCallback : Java.Lang.Object, IActivityResultCallback
    {
        public void OnActivityResult(Java.Lang.Object? result)
        {
            try
            {
                if (result == null)
                {
                    _tcs?.TrySetResult(string.Empty);
                    return;
                }

                var uri = result as AndroidUri;
                if (uri == null)
                {
                    _tcs?.TrySetResult(string.Empty);
                    return;
                }

                // 获取文件路径
                var path = GetFilePathFromUri(MainActivity.Current, uri);
                _tcs?.TrySetResult(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理文件选择结果失败: {ex.Message}");
                _tcs?.TrySetResult(string.Empty);
            }
        }

        /// <summary>
        /// 从 Uri 获取文件路径
        /// </summary>
        private string GetFilePathFromUri(Context? context, AndroidUri uri)
        {
            try
            {
                // 持久化 URI 权限
                var contentResolver = context?.ContentResolver;
                if (contentResolver == null)
                {
                    // 直接返回 Uri 字符串（需要特殊处理）
                    return uri.ToString();
                }

                // 尝试获取文件流
                var inputStream = contentResolver.OpenInputStream(uri);
                if (inputStream == null)
                {
                    return string.Empty;
                }

                // 获取文件名
                var cursor = contentResolver.Query(uri, null, null, null, null);
                string? fileName = null;
                if (cursor != null && cursor.MoveToFirst())
                {
                    int nameIndex = cursor.GetColumnIndex(IOpenableColumns.DisplayName);
                    if (nameIndex >= 0)
                    {
                        fileName = cursor.GetString(nameIndex);
                    }
                    cursor.Close();
                }

                // 复制文件到应用私有目录
                var tempDir = context?.GetExternalFilesDir(null)?.AbsolutePath;
                if (string.IsNullOrEmpty(tempDir) || string.IsNullOrEmpty(fileName))
                {
                    return string.Empty;
                }

                var tempPath = System.IO.Path.Combine(tempDir, fileName);
                using (var outputStream = System.IO.File.Create(tempPath))
                {
                    inputStream.CopyTo(outputStream);
                }
                inputStream.Close();

                return tempPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取文件路径失败: {ex.Message}");
                return string.Empty;
            }
        }
    }
}
