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
        System.Diagnostics.Debug.WriteLine("========== AndroidFilePickerService 静态构造函数 ==========");
        try
        {
            // 获取 MainActivity 实例
            var activity = MainActivity.Current;
            if (activity == null)
            {
                System.Diagnostics.Debug.WriteLine("❌ 无法获取 MainActivity 实例");
                return;
            }

            System.Diagnostics.Debug.WriteLine("✅ MainActivity 实例获取成功");

            // 注册文件选择器
            _pickerLauncher = activity.RegisterForActivityResult(
                new ActivityResultContracts.GetContent(),
                new ActivityResultCallback()
            );

            System.Diagnostics.Debug.WriteLine("✅ 文件选择器注册成功");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"❌ 初始化文件选择器失败: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"   异常类型: {ex.GetType().Name}");
        }
    }

    /// <summary>
    /// 选择 Excel 文件
    /// </summary>
    public async Task<string> PickExcelFileAsync()
    {
        System.Diagnostics.Debug.WriteLine("========== PickExcelFileAsync 开始 ==========");
        _tcs = new TaskCompletionSource<string>();

        try
        {
            var activity = MainActivity.Current;
            if (activity == null)
            {
                System.Diagnostics.Debug.WriteLine("❌ 无法获取 MainActivity 实例");
                return string.Empty;
            }

            System.Diagnostics.Debug.WriteLine("✅ MainActivity 实例有效");

            // GetContent() 直接传递 MIME 类型字符串
            string mimeType = "*/*";
            System.Diagnostics.Debug.WriteLine($"✅ 正在启动文件选择器，MIME 类型: {mimeType}");

            // 启动文件选择器
            _pickerLauncher.Launch(mimeType);

            System.Diagnostics.Debug.WriteLine("✅ 文件选择器已启动，等待用户选择");

            // 等待用户选择
            var result = await _tcs.Task;
            System.Diagnostics.Debug.WriteLine($"✅ 用户选择完成，结果: {result}");
            return result;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"❌ 文件选择失败: {ex.Message}");
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
            System.Diagnostics.Debug.WriteLine("========== OnActivityResult 被调用 ==========");
            try
            {
                if (result == null)
                {
                    System.Diagnostics.Debug.WriteLine("❌ result 为 null");
                    _tcs?.TrySetResult(string.Empty);
                    return;
                }

                var uri = result as AndroidUri;
                if (uri == null)
                {
                    System.Diagnostics.Debug.WriteLine("❌ uri 转换失败");
                    _tcs?.TrySetResult(string.Empty);
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"✅ URI 获取成功: {uri}");

                // 获取文件路径
                var path = GetFilePathFromUri(MainActivity.Current, uri);
                System.Diagnostics.Debug.WriteLine($"✅ 文件路径结果: {path}");
                _tcs?.TrySetResult(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ 处理文件选择结果失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"   异常类型: {ex.GetType().Name}");
                _tcs?.TrySetResult(string.Empty);
            }
        }

        /// <summary>
        /// 从 Uri 获取文件路径
        /// </summary>
        private string GetFilePathFromUri(Context? context, AndroidUri uri)
        {
            System.Diagnostics.Debug.WriteLine($"========== GetFilePathFromUri 开始 ==========");
            System.Diagnostics.Debug.WriteLine($"URI: {uri}");
            System.Diagnostics.Debug.WriteLine($"Context: {(context != null ? "有效" : "null")}");

            try
            {
                // 持久化 URI 权限
                var contentResolver = context?.ContentResolver;
                if (contentResolver == null)
                {
                    System.Diagnostics.Debug.WriteLine("❌ ContentResolver 为 null");
                    return string.Empty;
                }

                // 持久化读取权限
                context.GrantUriPermission(context.PackageName, uri, global::Android.Content.ActivityFlags.GrantReadUriPermission);

                // 尝试获取文件流
                var inputStream = contentResolver.OpenInputStream(uri);
                if (inputStream == null)
                {
                    System.Diagnostics.Debug.WriteLine("❌ 无法打开输入流");
                    return string.Empty;
                }

                System.Diagnostics.Debug.WriteLine("✅ 输入流打开成功");

                // 获取文件名
                var cursor = contentResolver.Query(uri, null, null, null, null);
                string? fileName = null;
                if (cursor != null && cursor.MoveToFirst())
                {
                    int nameIndex = cursor.GetColumnIndex(IOpenableColumns.DisplayName);
                    if (nameIndex >= 0)
                    {
                        fileName = cursor.GetString(nameIndex);
                        System.Diagnostics.Debug.WriteLine($"✅ 文件名: {fileName}");
                    }
                    cursor.Close();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("⚠️ 无法获取文件名，使用默认名称");
                    fileName = "temp_excel.xlsx";
                }

                // 复制文件到应用私有目录
                var tempDir = context?.GetExternalFilesDir(null)?.AbsolutePath;
                if (string.IsNullOrEmpty(tempDir))
                {
                    System.Diagnostics.Debug.WriteLine("❌ 无法获取外部文件目录");
                    inputStream.Close();
                    return string.Empty;
                }

                var tempPath = System.IO.Path.Combine(tempDir, fileName);
                System.Diagnostics.Debug.WriteLine($"临时文件路径: {tempPath}");

                // 检查文件是否已存在
                if (System.IO.File.Exists(tempPath))
                {
                    System.Diagnostics.Debug.WriteLine($"✅ 文件已存在，直接使用: {tempPath}");
                    System.Diagnostics.Debug.WriteLine($"文件大小: {new System.IO.FileInfo(tempPath).Length} 字节");
                    inputStream.Close();
                    return tempPath;
                }

                System.Diagnostics.Debug.WriteLine("⚠️ 文件不存在，开始复制");

                // 文件不存在，则复制文件
                using (var outputStream = System.IO.File.Create(tempPath))
                {
                    inputStream.CopyTo(outputStream);
                    outputStream.Flush();
                }
                inputStream.Close();

                System.Diagnostics.Debug.WriteLine($"✅ 文件复制成功，大小: {new System.IO.FileInfo(tempPath).Length} 字节");
                return tempPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ 获取文件路径失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"   异常类型: {ex.GetType().Name}");
                return string.Empty;
            }
        }
    }
}
