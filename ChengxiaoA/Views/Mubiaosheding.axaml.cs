using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using ChengxiaoA.Models;
using System;
using System.Diagnostics;

namespace ChengxiaoA.Views;

/// <summary>
/// 目标设定窗口
/// </summary>
public partial class Mubiaosheding : Window
{
    public bool Isxinshezhi { get; private set; } = false;
    public static MubaioShuxin? shuxin1 { get; set; }

    public Mubiaosheding()
    {
        InitializeComponent();
    }

    private void Newmubiao_Click(object? sender, RoutedEventArgs e)
    {
        Isxinshezhi = true;
    }

    private async void OK_Click(object? sender, RoutedEventArgs e)
    {
        if (Isxinshezhi)
        {
            if (shuxin1 == null)
            {
                shuxin1 = new MubaioShuxin();
            }

            shuxin1.Name = NameTextBox?.Text ?? string.Empty;
            shuxin1.Time = TimeTextBox?.Text ?? string.Empty;
            shuxin1.Dangqianchengxiao = 0;

            if (double.TryParse(BiliTextBox?.Text, out double bili))
            {
                shuxin1.Bili = bili;
            }

            if (double.TryParse(TimeTextBox?.Text, out double timeHours))
            {
                shuxin1.Mubiaochengxiao = shuxin1.Bili * 10.5 * timeHours * 6;
            }

            // 获取存储提供程序
            var topLevel = TopLevel.GetTopLevel(this);
            var storageProvider = topLevel?.StorageProvider;

            if (storageProvider == null)
            {
                Debug.WriteLine("无法获取存储提供程序");
                return;
            }

            // 显示保存文件对话框
            var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "保存",
                DefaultExtension = "dat",
                SuggestedFileName = shuxin1.Name,
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("DAT文件")
                    {
                        Patterns = new[] { "*.dat" }
                    },
                    new FilePickerFileType("所有文件")
                    {
                        Patterns = new[] { "*.*" }
                    }
                }
            });

            // 如果用户点击了"保存"
            if (file != null)
            {
                try
                {
                    // 获取用户指定的文件路径
                    string filePath = file.Path.OriginalString;

                    // 保存到文件
                    DataFileHandler.SaveDataToFile(shuxin1, filePath);

                    Debug.WriteLine($"目标文件已保存到: {filePath}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"保存失败: {ex.Message}");
                    return;
                }
            }
            else
            {
                Debug.WriteLine("用户取消了保存");
                return;
            }

            Close();

            // 创建并显示目标窗口
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                var newWindow = new Zimubiao();
                newWindow.Show();
            });
        }
        Isxinshezhi = false;
    }

    private void Cancel_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }

    private async void Open_Click(object? sender, RoutedEventArgs e)
    {
        // 获取存储提供程序
        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider == null)
        {
            Debug.WriteLine("无法获取存储提供程序");
            return;
        }

        // 显示打开文件对话框
        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "打开",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("DAT文件")
                {
                    Patterns = new[] { "*.dat" }
                },
                new FilePickerFileType("所有文件")
                {
                    Patterns = new[] { "*.*" }
                }
            }
        });

        // 如果用户点击了"打开"
        if (files.Count > 0)
        {
            try
            {
                // 获取选择的文件路径
                string filePath = files[0].Path.OriginalString;

                // 加载数据
                shuxin1 = DataFileHandler.LoadDataFromFile<MubaioShuxin>(filePath);

                Debug.WriteLine($"目标文件已加载: {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载失败: {ex.Message}");
                return;
            }

            Close();

            // 创建并显示目标窗口
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                var newWindow = new Zimubiao();
                newWindow.Show();
            });
        }
    }
}
