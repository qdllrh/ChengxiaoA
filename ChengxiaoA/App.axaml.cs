using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core;
using Avalonia.Data.Core.Plugins;
using System.Linq;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using ChengxiaoA.ViewModels;
using ChengxiaoA.Views;
using System.Diagnostics;

namespace ChengxiaoA;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        // 根据平台设置全局默认字体
        FontFamily fontFamily;

#if ANDROID
        // Android 平台：尝试多个字体路径
        Debug.WriteLine("=== Android 字体调试 ===");

        // 尝试嵌入字体（通过 Avalonia 资源）
        try
        {
            fontFamily = new FontFamily("avares://ChengxiaoA/Assets/Fonts/#Source Han Sans CN");
            Debug.WriteLine($"尝试使用嵌入字体: {fontFamily}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"嵌入字体加载失败: {ex.Message}");
            // 回退到 Android 系统字体
            fontFamily = new FontFamily("sans-serif");
            Debug.WriteLine($"使用系统默认字体: {fontFamily}");
        }
#else
        // 其他平台：使用字体回退机制
        fontFamily = new FontFamily("SimHei,Microsoft YaHei,Arial,sans-serif");
#endif

        Debug.WriteLine($"最终使用字体: {fontFamily}");

        // 设置全局字体资源
        Application.Current!.Resources["CustomChineseFont"] = fontFamily;

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // Avoid duplicate validations from both Avalonia and the CommunityToolkit.
            // More info: https://docs.avaloniaui.net/docs/guides/development-guides/data-validation#manage-validationplugins
            DisableAvaloniaDataAnnotationValidation();
            desktop.MainWindow = new MainWindow
            {
                DataContext = new MainViewModel()
            };
        }
        else if (ApplicationLifetime is ISingleViewApplicationLifetime singleViewPlatform)
        {
            singleViewPlatform.MainView = new MainView
            {
                DataContext = new MainViewModel()
            };
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void DisableAvaloniaDataAnnotationValidation()
    {
        // Get an array of plugins to remove
        var dataValidationPluginsToRemove =
            BindingPlugins.DataValidators.OfType<DataAnnotationsValidationPlugin>().ToArray();

        // remove each entry found
        foreach (var plugin in dataValidationPluginsToRemove)
        {
            BindingPlugins.DataValidators.Remove(plugin);
        }
    }
}