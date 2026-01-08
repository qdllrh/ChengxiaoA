using Android.App;
using Android.Content.PM;
using Android.OS;
using Avalonia;
using Avalonia.Android;
using ChengxiaoA.Android.Services;
using ChengxiaoA.Services;

namespace ChengxiaoA.Android;

[Activity(
    Label = "ChengxiaoA.Android",
    Theme = "@style/MyTheme.NoActionBar",
    Icon = "@drawable/icon",
    MainLauncher = true,
    ConfigurationChanges = ConfigChanges.Orientation | ConfigChanges.ScreenSize | ConfigChanges.UiMode)]
public class MainActivity : AvaloniaMainActivity<App>
{
    private static MainActivity? _current;

    public static MainActivity? Current => _current;

    protected override AppBuilder CustomizeAppBuilder(AppBuilder builder)
    {
        return base.CustomizeAppBuilder(builder)
            .WithInterFont();
    }

    protected override void OnCreate(Bundle? savedInstanceState)
    {
        System.Diagnostics.Debug.WriteLine("========== MainActivity.OnCreate ==========");
        _current = this;
        base.OnCreate(savedInstanceState);

        // 注册 Android 文件选择器服务
        System.Diagnostics.Debug.WriteLine("正在注册 AndroidFilePickerService...");
        FilePickerService.Instance = new AndroidFilePickerService();
        System.Diagnostics.Debug.WriteLine("✅ AndroidFilePickerService 注册完成");
    }
}
