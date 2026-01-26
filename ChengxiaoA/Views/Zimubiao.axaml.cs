using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ChengxiaoA.Models;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ChengxiaoA.Views;

/// <summary>
/// 目标显示窗口
/// </summary>
public partial class Zimubiao : Window
{
    public static double MubiaoYutime { get; set; }
    public static double MubiaoDangqian { get; set; }
    public static string? Namezi { get; set; }

    private double _waveOffset = 0;
    private bool _isWaveAnimating = false;

    // 控件引用 - 只缓存需要手动获取的控件
    private Canvas? waveCanvasControl;

    public Zimubiao()
    {
        InitializeComponent();

        // 获取波纹画布（需要手动获取用于动画）
        waveCanvasControl = this.FindControl<Canvas>("waveCanvas");

        // 订阅窗口关闭事件
        this.Closed += OnWindowClosed;

        // 初始化数据
        if (Mubiaosheding.shuxin1 != null)
        {
            YujiTime!.Text = Mubiaosheding.shuxin1.Time;
            var cx = Mubiaosheding.shuxin1.Mubiaochengxiao;
            YujiTimeCX!.Text = cx.ToString();
            MubiaoDangqian = Mubiaosheding.shuxin1.Dangqianchengxiao;
            Dancichengxiaoleiji!.Text = Mubiaosheding.shuxin1.Dangqianchengxiao.ToString();
            Nameizi1!.Text = Mubiaosheding.shuxin1.Name;
            Title = Mubiaosheding.shuxin1.Name;
            TitleTextBlockControl!.Text = Mubiaosheding.shuxin1.Name;
        }

        UpdateWaterFill();
        StartWaveAnimation();
    }

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        _isWaveAnimating = false;
    }

    private void YuleJisuan_Click(object? sender, RoutedEventArgs e)
    {
        if (double.TryParse(YuleTime?.Text, out double yuleTime))
        {
            MubiaoYutime = yuleTime;
            MainView.total2 = MainView.total2 + MubiaoYutime;
            UpdateWaterFill();
            YuleTime!.Clear();
        }
    }

    private void Jisuan_Click(object? sender, RoutedEventArgs e)
    {
        if (Mubiaosheding.shuxin1 == null) return;

        if (double.TryParse(XuexiTime?.Text, out double xuexiMinutes))
        {
            var xuexiChengxiao = xuexiMinutes * 10.5 / 10 * Mubiaosheding.shuxin1.Bili;
            Dancichengxiao!.Text = xuexiChengxiao.ToString("F2");
            leiji1_1!.IsEnabled = true;
            jisuan1_1!.IsEnabled = false;
            XuexiTime!.Clear();
        }
    }

    private void Leiji_Click(object? sender, RoutedEventArgs e)
    {
        if (double.TryParse(Dancichengxiao?.Text, out double danciChengxiao))
        {
            MubiaoDangqian = MubiaoDangqian + danciChengxiao;
            Dancichengxiaoleiji!.Text = MubiaoDangqian.ToString();
            MainView.total = MainView.total + danciChengxiao;
            UpdateWaterFill();
            leiji1_1!.IsEnabled = false;
            jisuan1_1!.IsEnabled = true;
            MainView.Zongjilei1 = MainView.Zongjilei1 + danciChengxiao;
        }
    }

    private async void Save_Click(object? sender, RoutedEventArgs e)
    {
        if (Mubiaosheding.shuxin1 == null) return;

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
            SuggestedFileName = Mubiaosheding.shuxin1.Name,
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

                // 更新当前成效
                Mubiaosheding.shuxin1.Dangqianchengxiao = MubiaoDangqian;

                // 保存到文件
                DataFileHandler.SaveDataToFile(Mubiaosheding.shuxin1, filePath);

                Debug.WriteLine($"目标文件已保存到: {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存失败: {ex.Message}");
            }
        }
    }

    private void UpdateWaterFill()
    {
        if (Mubiaosheding.shuxin1 == null) return;

        // 计算高度比例
        double maxValue = Mubiaosheding.shuxin1.Mubiaochengxiao;
        if (maxValue <= 0) maxValue = 1;

        double percentage = Math.Min(MubiaoDangqian / maxValue, 1.0);
        double containerHeight = 182.24;
        double height = containerHeight * percentage;

        // 更新填充高度
        waterFill!.Height = height;
        UpdateWaterColor();

        Debug.WriteLine($"水灌更新: {MubiaoDangqian}/{maxValue} = {percentage:P0}, 高度 = {height:F2}");
    }

    private void UpdateWaterColor()
    {
        Color topColor = Color.FromArgb(255, 100, 255, 100);
        Color bottomColor = Color.FromArgb(255, 0, 180, 0);

        // 更新渐变颜色
        if (waterFill!.Background is LinearGradientBrush brush)
        {
            brush.GradientStops[0].Color = topColor;
            brush.GradientStops[1].Color = bottomColor;
        }
    }

    private void StartWaveAnimation()
    {
        if (_isWaveAnimating) return;
        _isWaveAnimating = true;

        Task.Run(async () =>
        {
            while (_isWaveAnimating)
            {
                _waveOffset -= 1;
                if (_waveOffset <= -60) _waveOffset = 0;

                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    // 使用 RenderTransform 来移动波纹
                    var transform = new TranslateTransform(_waveOffset, 0);
                    waveCanvasControl!.RenderTransform = transform;
                });

                await Task.Delay(50);
            }
        });
    }

}
