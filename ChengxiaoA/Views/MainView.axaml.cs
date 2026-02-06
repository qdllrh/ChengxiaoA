using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform;
using Avalonia.Platform.Storage;
using ChengxiaoA.Services;
using ChengxiaoA.Models;
using System;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace ChengxiaoA.Views;

public partial class MainView : UserControl
{
    // 静态字段：用于存储累计成效
    public static double total = 0;
    public static double total2 = 0;
    public static double Xuexivalue;
    public static double Yulevalue;
    public static string _selectedExcelPath = string.Empty;
    public static string _selectedExcelPath1 = string.Empty;
    public static bool IsHuanzhai = false;
    public static double Zongjilei1;
    public static double Duoyuleji = 0;
    // 注入 Excel 读取服务（只依赖接口，不依赖 Android 实现）
    private readonly IExcelReaderService _excelReaderService;

    // 控件缓存（避免每次调用 FindControl，提升性能）
    private Grid? _waterContainer;
    private Border? _waterFill;
    private Grid? _waterContainer1;
    private Border? _waterFill1;
    private Grid? _waterContainer2;
    private Border? _waterFill2;

    // 目标设定相关
    private static MubaioShuxin? _mubiaoshuxin = null;
    private static double _mubiaoDangqian = 0;
    private static double _mubiaoYutime = 0;
    private static double _waveOffsetZimubiao = 0;
    private static bool _isWaveAnimatingZimubiao = false;

    public MainView()
    {
        InitializeComponent();

        // 注册 Loaded 事件
        this.Loaded += MainWindow_Loaded;

        // 初始化水灌动画
        InitializeAnimations();

        // 初始化控件缓存
        CacheWaterControls();

    }

   
    // 初始化水灌控件缓存
    private void CacheWaterControls()
    {
        _waterContainer = this.FindControl<Grid>("waterContainer");
        _waterFill = this.FindControl<Border>("waterFill");
        _waterContainer1 = this.FindControl<Grid>("waterContainer1");
        _waterFill1 = this.FindControl<Border>("waterFill1");
        _waterContainer2 = this.FindControl<Grid>("waterContainer2");
        _waterFill2 = this.FindControl<Border>("waterFill2");
    }

    #region 窗口加载和关闭事件

    // 窗口加载事件 - 读取 Excel 文件
    private async void MainWindow_Loaded(object? sender, RoutedEventArgs e)
    {
        try
        {
            // 调用平台特定的文件选择器
            string excelFilePath = await PickExcelFileAsync();

            if (!string.IsNullOrEmpty(excelFilePath))
                {
                var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
                if (txtXuexizongjilei != null)
                {
                    txtXuexizongjilei.Text = "正在读取文件...";

                    // 读取第一个单元格的数字
                    double firstCellNumber;

                    // 运行时检测是否在 Android 平台（使用多种检测方法）
                    bool isAndroid = false;

                    // 方法1: 检测 RuntimeIdentifier（最可靠）
                    string runtimeId = System.Runtime.InteropServices.RuntimeInformation.RuntimeIdentifier ?? "";
                    isAndroid = runtimeId.Contains("android", StringComparison.OrdinalIgnoreCase);

                    Debug.WriteLine($"=== 平台检测 ===");
                    Debug.WriteLine($"运行时信息: {runtimeId}");
                    Debug.WriteLine($"isAndroid (基于 RuntimeIdentifier): {isAndroid}");
                    Debug.WriteLine($"操作系统: {System.Environment.OSVersion}");

                    if (isAndroid)
                    {
                        // Android 平台：使用服务容器（通过反射避免编译时依赖）
                        Debug.WriteLine("检测到 Android 平台，尝试获取服务容器...");

                        // 尝试从已加载的程序集中查找类型
                        Type? containerType = null;

                        // 注意：AndroidServiceContainer 是 AndroidExcelReaderService 的嵌套类
                        // 正确的类型名应该是：ChengxiaoA.Android.Services.AndroidExcelReaderService+AndroidServiceContainer
                        const string fullTypeName = "ChengxiaoA.Android.Services.AndroidExcelReaderService+AndroidServiceContainer";

                        // 尝试遍历所有程序集
                        foreach (var assembly in System.AppDomain.CurrentDomain.GetAssemblies())
                        {
                            if (assembly.GetName().Name?.Contains("ChengxiaoA.Android", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                Debug.WriteLine($"找到 Android 程序集: {assembly.GetName().Name}");
                                containerType = assembly.GetType(fullTypeName);
                                if (containerType != null)
                                {
                                    Debug.WriteLine($"  找到类型: {containerType.FullName}");
                                    break;
                                }
                            }
                        }

                        Debug.WriteLine($"containerType: {containerType?.FullName ?? "null"}");
                        if (containerType != null)
                        {
                            var serviceProperty = containerType.GetProperty("ExcelReaderService", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                            Debug.WriteLine($"serviceProperty: {serviceProperty?.Name ?? "null"}");
                            var reader = serviceProperty?.GetValue(null) as IExcelReaderService;
                            Debug.WriteLine($"reader: {reader?.GetType().Name ?? "null"}");
                            if (reader != null)
                            {
                                firstCellNumber = reader.ReadFirstCellNumber(excelFilePath);
                            }
                            else
                            {
                                Debug.WriteLine("reader 为空，使用本地方法");
                                firstCellNumber = ReadFirstCellNumber(excelFilePath);
                            }
                        }
                        else
                        {
                            Debug.WriteLine("服务容器类型为空，使用本地方法");
                            firstCellNumber = ReadFirstCellNumber(excelFilePath);
                        }
                    }
                    else
                    {
                        // 其他平台：直接使用本地方法
                        firstCellNumber = ReadFirstCellNumber(excelFilePath);
                    }

                    // 显示结果
                    txtXuexizongjilei.Text = $"{firstCellNumber}";
                    Zongjilei1 = firstCellNumber;
                    _selectedExcelPath1 = excelFilePath;
                }
            }
            else
            {
                // 文件不存在，显示提示
                var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
                if (txtXuexizongjilei != null)
                {
                    txtXuexizongjilei.Text = "请选择 Excel 文件";
                }
            }
        }
        catch (Exception ex)
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
            if (txtXuexizongjilei != null)
            {
                txtXuexizongjilei.Text = "读取失败，请检查文件格式";
            }
        }
//#endif
    }

    // 窗口关闭事件 - 保存到 Excel
    private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            if (string.IsNullOrEmpty(_selectedExcelPath1))
            {
                return;
            }

            // 同步保存
            SaveValueToFirstCell(_selectedExcelPath1, Zongjilei1);
        }
        catch (Exception ex)
        {
        }
    }

    // 获取默认 Excel 路径（Android 需要特殊处理）
    private string GetDefaultExcelPath()
    {
        // TODO: 根据平台返回不同的路径
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "chengxiao_data.xlsx");
    }

    #endregion

    #region 计算成效按钮点击事件

    private void BtnCalculate_Click(object? sender, RoutedEventArgs e)
    {
        var txtX = this.FindControl<TextBox>("txtX");
        var txtY = this.FindControl<TextBox>("txtY");
        var txtZ = this.FindControl<TextBox>("txtZ");
        var txtResult = this.FindControl<TextBox>("txtResult");
        var btnCalculate = this.FindControl<Button>("btnCalculate");
        var btnAccumulate = this.FindControl<Button>("btnAccumulate");

        if (txtX != null && txtY != null && txtZ != null && txtResult != null)
        {
            if (double.TryParse(txtX.Text, out double x) &&
                double.TryParse(txtY.Text, out double y) &&
                double.TryParse(txtZ.Text, out double z))
            {
                // 确保Y不为0（避免除数为0）
                if (y == 0)
                {
                    return;
                }

                // 按照公式计算成效：|(x+2y) * (x/(1.5y) * 0.5 + 0.5 * z)|
                double part1 = x + 2 * y;
                double part2;
                double aa = 0.75;

                if (x <= 1.5 * y)
                {
                    part2 = (x / (1.5 * y)) * 0.5 + 0.5 * z;
                    if (x / (1.5 * y) < 0.2)
                    {
                        part2 = 0.2 * 0.5 + 0.5 * z;
                    }
                }
                else
                {
                    part2 = ((1.5 * y) / x) * 0.5 + 0.5 * z;
                }

                double result = Math.Abs(part1 * part2);
                txtResult.Text = result.ToString("F2"); // 保留2位小数

                // 清空XYZ输入框
                txtX.Clear();
                txtY.Clear();
                txtZ.Clear();

                // 暂时禁用计算按钮，启用积累按钮
                if (btnCalculate != null)
                    btnCalculate.IsEnabled = false;
                if (btnAccumulate != null)
                    btnAccumulate.IsEnabled = true;
            }
        }
    }

    #endregion

    #region 累计按钮点击事件

    private void BtnAccumulate_Click(object? sender, RoutedEventArgs e)
    {
        var txtResult = this.FindControl<TextBox>("txtResult");
        var txtTotal = this.FindControl<TextBox>("txtTotal");
        var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
        var btnCalculate = this.FindControl<Button>("btnCalculate");
        var btnAccumulate = this.FindControl<Button>("btnAccumulate");

        if (txtResult != null)
        {
            if (double.TryParse(txtResult.Text, out double currentResult))
            {
                total += currentResult; // 累加当前成效
                Zongjilei1 += currentResult;

                if (txtTotal != null)
                    txtTotal.Text = total.ToString("F2"); // 显示累计结果

                if (txtXuexizongjilei != null)
                    txtXuexizongjilei.Text = Zongjilei1.ToString("F2");

                UpdateWaterFill(total, 0); // 更新水灌效果
                UpdateWaterFill(Zongjilei1, 1); // 更新水灌效果
                Xuexivalue = total;
                UpdateWaterColor2(total2);

                // 重新启用计算按钮，禁用积累按钮
                if (btnCalculate != null)
                    btnCalculate.IsEnabled = true;
                if (btnAccumulate != null)
                    btnAccumulate.IsEnabled = false;
            }
        }
    }

    #endregion

    #region 娱乐相关事件

    private void Yulebutton_Click(object? sender, RoutedEventArgs e)
    {
        var txtYleshijian = this.FindControl<TextBox>("txtYleshijian");
        var txtYule = this.FindControl<TextBox>("txtYule");
        var Yulebutton = this.FindControl<Button>("Yulebutton");
        var Yuleleijibutton = this.FindControl<Button>("Yuleleijibutton");

        if (txtYleshijian != null && txtYule != null)
        {
            if (double.TryParse(txtYleshijian.Text, out double yule))
            {
                double resultyule = yule / 10 * 10.5 * 1.75;
                txtYule.Text = resultyule.ToString("F2"); // 保留2位小数
                txtYleshijian.Clear();

                // 禁用娱乐计算按钮，启用娱乐积累按钮
                if (Yulebutton != null)
                    Yulebutton.IsEnabled = false;
                if (Yuleleijibutton != null)
                    Yuleleijibutton.IsEnabled = true;
            }
        }
    }

    private void Yuleleijibutton_Click(object? sender, RoutedEventArgs e)
    {
        var txtYule = this.FindControl<TextBox>("txtYule");
        var txtleijiYule = this.FindControl<TextBox>("txtleijiYule");
        var Yulebutton = this.FindControl<Button>("Yulebutton");
        var Yuleleijibutton = this.FindControl<Button>("Yuleleijibutton");

        if (txtYule != null && txtleijiYule != null)
        {
            if (double.TryParse(txtYule.Text, out double currentResult))
            {
                total2 += currentResult; // 累加当前成效
                txtleijiYule.Text = total2.ToString("F2"); // 显示累计结果
                UpdateWaterFill(total2, 2); // 更新水灌效果
                Yulevalue = total2;

                // 重新启用娱乐计算按钮，禁用娱乐积累按钮
                if (Yulebutton != null)
                    Yulebutton.IsEnabled = true;
                if (Yuleleijibutton != null)
                    Yuleleijibutton.IsEnabled = false;
            }
        }
    }

    #endregion



    #region 总计相关事件

    private async void Zongjilei_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
            if (txtXuexizongjilei != null)
                txtXuexizongjilei.Text = "请选择Excel文件...";

            // 调用平台特定的文件选择器
            string excelFilePath = await PickExcelFileAsync();

            if (!string.IsNullOrEmpty(excelFilePath) && File.Exists(excelFilePath))
            {
                if (txtXuexizongjilei != null)
                    txtXuexizongjilei.Text = "正在读取文件...";

                // 读取第一个单元格的数字
                double firstCellNumber = ReadFirstCellNumber(excelFilePath);

                if (txtXuexizongjilei != null)
                    txtXuexizongjilei.Text = $"{firstCellNumber}";

                _selectedExcelPath = excelFilePath;
            }
            else if (txtXuexizongjilei != null)
            {
                txtXuexizongjilei.Text = "未选择文件";
            }
        }
        catch (Exception ex)
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
            if (txtXuexizongjilei != null)
            {
                txtXuexizongjilei.Text = "读取失败，请检查文件格式";
            }
        }
    }

    // 平台特定的文件选择器
    private async Task<string> PickExcelFileAsync()
    {
        // 运行时判断平台 - 尝试多种检测方式
        bool isAndroid = false;

        // 方法1: 检查 Android 类型
        try
        {
            isAndroid = System.Type.GetType("Android.App.Activity", false) != null;
        }
        catch { }

        // 方法2: 如果方法1失败，检查操作系统描述符
        if (!isAndroid)
        {
            try
            {
                var osPlatform = System.Runtime.InteropServices.RuntimeInformation.OSDescription;
                isAndroid = osPlatform.Contains("Android", System.StringComparison.OrdinalIgnoreCase);
            }
            catch { }
        }

        // 方法3: 检查应用是否运行在 Mono.Android 环境
        if (!isAndroid)
        {
            try
            {
                isAndroid = System.Type.GetType("Mono.Android.Runtime", false) != null;
            }
            catch { }
        }

        if (isAndroid)
        {
            // Android 平台：使用存储访问框架
            return await FilePickerService.PickExcelFileAsync();
        }
        else
        {
            // 其他平台：使用存储文件对话框
            // 获取 TopLevel 对象
            var topLevel = TopLevel.GetTopLevel(this);
            var storageProvider = topLevel?.StorageProvider;

            if (storageProvider == null)
                return string.Empty;

            var file = await storageProvider.OpenFilePickerAsync(new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title = "选择Excel文件",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel文件")
                    {
                        Patterns = new[] { "*.xls", "*.xlsx" },
                        AppleUniformTypeIdentifiers = new[] { "com.microsoft.excel.xls" },
                        MimeTypes = new[] { "application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                    },
                    new Avalonia.Platform.Storage.FilePickerFileType("所有文件")
                    {
                        Patterns = new[] { "*.*" }
                    }
                }
            });

            // Avalonia 11.x 中使用 Path.AbsolutePath 而不是 TryGetLocalPath()
            return file.Count > 0 ? file[0].Path.AbsolutePath : string.Empty;
        }
    }

    private void ZongjileiJisuan_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");

            if (string.IsNullOrEmpty(_selectedExcelPath))
            {
                return;
            }

            double.TryParse(txtXuexizongjilei?.Text, out double valueToSave);

            SaveValueToFirstCell(_selectedExcelPath, valueToSave);

            if (txtXuexizongjilei != null)
                txtXuexizongjilei.Text = $"{valueToSave}";
        }
        catch (Exception ex)
        {
        }
    }

    #endregion

    #region 保存和加载事件

    private async void Baocun_Click(object? sender, RoutedEventArgs e)
    {
        // 获取 TopLevel 对象
        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider == null)
        {
            Debug.WriteLine("无法获取存储提供程序");
            return;
        }

        // 显示保存文件对话框
        var file = await storageProvider.SaveFilePickerAsync(new Avalonia.Platform.Storage.FilePickerSaveOptions
        {
            Title = "保存",
            DefaultExtension = "dat",
            SuggestedFileName = "chengxiao_data.dat",
            FileTypeChoices = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("DAT文件")
                {
                    Patterns = new[] { "*.dat" }
                },
                new Avalonia.Platform.Storage.FilePickerFileType("所有文件")
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
                // 在 Android 上，OriginalString 包含完整的 content:// URI
                // AbsolutePath 只包含 /document/xx 这样的简化路径
                string filePath = file.Path.OriginalString;
                
                Debug.WriteLine($"=== Baocun_Click 开始 ===");
                Debug.WriteLine($"文件对象: {file}");
                Debug.WriteLine($"文件路径 (AbsolutePath): {file.Path.AbsolutePath}");
                Debug.WriteLine($"文件路径 (OriginalString): {filePath}");
                Debug.WriteLine($"文件路径 (LocalPath): {file.Path.LocalPath}");
                Debug.WriteLine($"文件 URI: {file.Path}");

                // 创建数据对象
                ChengxiaoShuxin chengxiaoshuxin = new ChengxiaoShuxin
                {
                    Chengxiaoleiji = total,
                    Yuleleiji = total2,
                    Duoyuleiji = Duoyuleji
                };

                Debug.WriteLine($"准备保存的数据: Chengxiaoleiji={chengxiaoshuxin.Chengxiaoleiji}, Yuleleiji={chengxiaoshuxin.Yuleleiji}, Duoyuleiji={chengxiaoshuxin.Duoyuleiji}");

                // 保存到文件
                DataFileHandler.SaveDataToFile(chengxiaoshuxin, filePath);

                Debug.WriteLine($"文件已保存到: {filePath}");
                Debug.WriteLine($"=== Baocun_Click 结束 ===");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存失败: {ex.Message}");
                Debug.WriteLine($"异常类型: {ex.GetType().Name}");
                Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
            }
        }
    }

    private async void Jiazai_Click(object? sender, RoutedEventArgs e)
    {
        // 获取 TopLevel 对象
        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider == null)
        {
            Debug.WriteLine("无法获取存储提供程序");
            return;
        }

        // 显示打开文件对话框
        var files = await storageProvider.OpenFilePickerAsync(new Avalonia.Platform.Storage.FilePickerOpenOptions
        {
            Title = "打开",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("DAT文件")
                {
                    Patterns = new[] { "*.dat" }
                },
                new Avalonia.Platform.Storage.FilePickerFileType("所有文件")
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
                // 在 Android 上，OriginalString 包含完整的 content:// URI
                string filePath = files[0].Path.OriginalString;

                // 加载数据
                ChengxiaoShuxin gongyongShuxing1 = DataFileHandler.LoadDataFromFile<ChengxiaoShuxin>(filePath);
                total = gongyongShuxing1.Chengxiaoleiji;
                total2 = gongyongShuxing1.Yuleleiji;
                Duoyuleji = gongyongShuxing1.Duoyuleiji;
                Jiazaixiaoguo();

                Debug.WriteLine($"文件已加载: {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载失败: {ex.Message}");
            }
        }
    }

    private void Jiazaixiaoguo()
    {
        var txtTotal = this.FindControl<TextBox>("txtTotal");
        var txtleijiYule = this.FindControl<TextBox>("txtleijiYule");
        var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");

        if (txtTotal != null)
            txtTotal.Text = total.ToString("F2");

        UpdateWaterFill(total, 0);

        if (txtleijiYule != null)
            txtleijiYule.Text = total2.ToString("F2");

        UpdateWaterFill(total2, 2);

        UpdateWaterFill(Zongjilei1, 1);

        if (txtXuexizongjilei != null)
            txtXuexizongjilei.Text = Zongjilei1.ToString("F2");
    }

    #endregion

    #region 其他按钮事件

    private void Mubiaodakai_Click(object? sender, RoutedEventArgs e)
    {
        // 显示目标设定面板
        var mubiaoPanel = this.FindControl<Border>("MubiaoPanel");
        if (mubiaoPanel != null)
        {
            mubiaoPanel.IsVisible = true;
        }
    }

    // 目标设定：打开
    private async void MubiaoOpen_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider == null) return;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "打开",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("DAT文件") { Patterns = new[] { "*.dat" } },
                new FilePickerFileType("所有文件") { Patterns = new[] { "*.*" } }
            }
        });

        if (files.Count > 0)
        {
            try
            {
                _mubiaoshuxin = DataFileHandler.LoadDataFromFile<MubaioShuxin>(files[0].Path.OriginalString);
                ShowZimubiaoPanel();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载失败: {ex.Message}");
            }
        }
    }

    // 目标设定：新设置
    private void MubiaoNew_Click(object? sender, RoutedEventArgs e)
    {
        // 清空输入框
        var nameBox = this.FindControl<TextBox>("MubiaoNameTextBox");
        var timeBox = this.FindControl<TextBox>("MubiaoTimeTextBox");
        var biliBox = this.FindControl<TextBox>("MubiaoBiliTextBox");

        if (nameBox != null) nameBox.Clear();
        if (timeBox != null) timeBox.Clear();
        if (biliBox != null) biliBox.Clear();
    }

    // 目标设定：确定
    private async void MubiaoOK_Click(object? sender, RoutedEventArgs e)
    {
        var nameBox = this.FindControl<TextBox>("MubiaoNameTextBox");
        var timeBox = this.FindControl<TextBox>("MubiaoTimeTextBox");
        var biliBox = this.FindControl<TextBox>("MubiaoBiliTextBox");

        if (nameBox == null || timeBox == null || biliBox == null) return;

        if (_mubiaoshuxin == null)
        {
            _mubiaoshuxin = new MubaioShuxin();
        }

        _mubiaoshuxin.Name = nameBox.Text ?? string.Empty;
        _mubiaoshuxin.Time = timeBox.Text ?? string.Empty;
        _mubiaoshuxin.Dangqianchengxiao = 0;

        if (double.TryParse(biliBox.Text, out double bili))
        {
            _mubiaoshuxin.Bili = bili;
        }

        if (double.TryParse(timeBox.Text, out double timeHours))
        {
            _mubiaoshuxin.Mubiaochengxiao = _mubiaoshuxin.Bili * 10.5 * timeHours * 6;
        }

        // 保存文件
        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider != null)
        {
            var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "保存",
                DefaultExtension = "dat",
                SuggestedFileName = _mubiaoshuxin.Name,
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("DAT文件") { Patterns = new[] { "*.dat" } },
                    new FilePickerFileType("所有文件") { Patterns = new[] { "*.*" } }
                }
            });

            if (file != null)
            {
                try
                {
                    DataFileHandler.SaveDataToFile(_mubiaoshuxin, file.Path.OriginalString);
                    HideMubiaoPanel();
                    ShowZimubiaoPanel();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"保存失败: {ex.Message}");
                }
            }
        }
    }

    // 目标设定：取消
    private void MubiaoCancel_Click(object? sender, RoutedEventArgs e)
    {
        HideMubiaoPanel();
    }

    // 显示目标面板
    private void ShowZimubiaoPanel()
    {
        if (_mubiaoshuxin == null) return;

        // 隐藏目标设定面板
        HideMubiaoPanel();

        // 显示目标面板并填充数据
        var zimubiaoPanel = this.FindControl<Border>("ZimubiaoPanel");
        var nameBox = this.FindControl<TextBox>("ZimubiaoName");
        var yujiTime = this.FindControl<TextBox>("ZimubiaoYujiTime");
        var yujiCX = this.FindControl<TextBox>("ZimubiaoYujiCX");
        var dangqianCX = this.FindControl<TextBox>("ZimubiaoDangqianCX");

        if (zimubiaoPanel != null) zimubiaoPanel.IsVisible = true;
        if (nameBox != null) nameBox.Text = _mubiaoshuxin.Name;
        if (yujiTime != null) yujiTime.Text = _mubiaoshuxin.Time;
        if (yujiCX != null) yujiCX.Text = _mubiaoshuxin.Mubiaochengxiao.ToString();
        if (dangqianCX != null) dangqianCX.Text = _mubiaoshuxin.Dangqianchengxiao.ToString();

        _mubiaoDangqian = _mubiaoshuxin.Dangqianchengxiao;
        UpdateZimubiaoWaterFill();
        StartZimubiaoWaveAnimation();
    }

    // 隐藏目标设定面板
    private void HideMubiaoPanel()
    {
        var mubiaoPanel = this.FindControl<Border>("MubiaoPanel");
        if (mubiaoPanel != null) mubiaoPanel.IsVisible = false;
    }

    // 隐藏目标显示面板
    private void HideZimubiaoPanel()
    {
        var zimubiaoPanel = this.FindControl<Border>("ZimubiaoPanel");
        if (zimubiaoPanel != null) zimubiaoPanel.IsVisible = false;
        _isWaveAnimatingZimubiao = false;
    }

    // 目标：娱乐计算
    private void ZimubiaoYule_Click(object? sender, RoutedEventArgs e)
    {
        var yuleTimeBox = this.FindControl<TextBox>("ZimubiaoYuleTime");
        if (yuleTimeBox == null) return;

        if (double.TryParse(yuleTimeBox.Text, out double yuleTime))
        {
            _mubiaoYutime = yuleTime;
            total2 = total2 + _mubiaoYutime;
            UpdateWaterFill(total2, 2);
            yuleTimeBox.Clear();
        }
    }

    // 目标：学习计算
    private void ZimubiaoJisuan_Click(object? sender, RoutedEventArgs e)
    {
        if (_mubiaoshuxin == null) return;

        var xuexiTimeBox = this.FindControl<TextBox>("ZimubiaoXuexiTime");
        var danciCXBox = this.FindControl<TextBox>("ZimubiaoDanciCX");
        var leijiButton = this.FindControl<Button>("ZimubiaoLeijiButton");

        if (xuexiTimeBox == null || danciCXBox == null || leijiButton == null) return;

        if (double.TryParse(xuexiTimeBox.Text, out double xuexiMinutes))
        {
            var xuexiChengxiao = xuexiMinutes * 10.5 / 10 * _mubiaoshuxin.Bili;
            danciCXBox.Text = xuexiChengxiao.ToString("F2");
            leijiButton.IsEnabled = true;
            xuexiTimeBox.Clear();
        }
    }

    // 目标：累计
    private void ZimubiaoLeiji_Click(object? sender, RoutedEventArgs e)
    {
        var danciCXBox = this.FindControl<TextBox>("ZimubiaoDanciCX");
        var dangqianCXBox = this.FindControl<TextBox>("ZimubiaoDangqianCX");
        var leijiButton = this.FindControl<Button>("ZimubiaoLeijiButton");
        var jisuanButton = this.FindControl<Button>("ZimubiaoJisuanButton");

        if (danciCXBox == null || dangqianCXBox == null || leijiButton == null || jisuanButton == null) return;

        if (double.TryParse(danciCXBox.Text, out double danciChengxiao))
        {
            _mubiaoDangqian = _mubiaoDangqian + danciChengxiao;
            dangqianCXBox.Text = _mubiaoDangqian.ToString();
            total = total + danciChengxiao;
            UpdateZimubiaoWaterFill();
            leijiButton.IsEnabled = false;
            jisuanButton.IsEnabled = true;
            Zongjilei1 = Zongjilei1 + danciChengxiao;
        }
    }

    // 目标：保存
    private async void ZimubiaoSave_Click(object? sender, RoutedEventArgs e)
    {
        if (_mubiaoshuxin == null) return;

        var topLevel = TopLevel.GetTopLevel(this);
        var storageProvider = topLevel?.StorageProvider;

        if (storageProvider == null) return;

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "保存",
            DefaultExtension = "dat",
            SuggestedFileName = _mubiaoshuxin.Name,
            FileTypeChoices = new[]
            {
                new FilePickerFileType("DAT文件") { Patterns = new[] { "*.dat" } },
                new FilePickerFileType("所有文件") { Patterns = new[] { "*.*" } }
            }
        });

        if (file != null)
        {
            try
            {
                _mubiaoshuxin.Dangqianchengxiao = _mubiaoDangqian;
                DataFileHandler.SaveDataToFile(_mubiaoshuxin, file.Path.OriginalString);
                Debug.WriteLine($"目标文件已保存到: {file.Path.OriginalString}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存失败: {ex.Message}");
            }
        }
    }

    // 更新目标水灌效果
    private void UpdateZimubiaoWaterFill()
    {
        if (_mubiaoshuxin == null) return;

        double maxValue = _mubiaoshuxin.Mubiaochengxiao;
        if (maxValue <= 0) maxValue = 1;

        double percentage = Math.Min(_mubiaoDangqian / maxValue, 1.0);
        var waterContainer = this.FindControl<Grid>("waterContainerZimubiao");
        var waterFill = this.FindControl<Border>("waterFillZimubiao");

        if (waterContainer != null && waterFill != null)
        {
            double height = waterContainer.Bounds.Height * percentage;
            waterFill.Height = height;
            UpdateZimubiaoWaterColor();
        }

        Debug.WriteLine($"目标水灌更新: {_mubiaoDangqian}/{maxValue} = {percentage:P0}");
    }

    private void UpdateZimubiaoWaterColor()
    {
        var waterFill = this.FindControl<Border>("waterFillZimubiao");
        if (waterFill?.Background is LinearGradientBrush brush)
        {
            brush.GradientStops[0].Color = Color.FromArgb(255, 100, 255, 100);
            brush.GradientStops[1].Color = Color.FromArgb(255, 0, 180, 0);
        }
    }

    // 目标水灌波纹动画
    private void StartZimubiaoWaveAnimation()
    {
        if (_isWaveAnimatingZimubiao) return;
        _isWaveAnimatingZimubiao = true;

        Task.Run(async () =>
        {
            while (_isWaveAnimatingZimubiao)
            {
                _waveOffsetZimubiao -= 1;
                if (_waveOffsetZimubiao <= -60) _waveOffsetZimubiao = 0;

                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
                {
                    var waveCanvas = this.FindControl<Canvas>("waveCanvasZimubiao");
                    if (waveCanvas != null)
                    {
                        var transform = new TranslateTransform(_waveOffsetZimubiao, 0);
                        waveCanvas.RenderTransform = transform;
                    }
                });

                await Task.Delay(50);
            }
        });
    }

    private void Shuaxin_Click(object? sender, RoutedEventArgs e)
    {
        Jiazaixiaoguo();
    }

    private void Huanzhai_Click(object? sender, RoutedEventArgs e)
    {
        var Huanzhai = this.FindControl<Button>("Huanzhai");

        if (IsHuanzhai)
        {
            IsHuanzhai = false;
            if (Huanzhai != null)
                Huanzhai.Background = Avalonia.Media.Brushes.Yellow;
        }
        else
        {
            IsHuanzhai = true;
            if (Huanzhai != null)
                Huanzhai.Background = Avalonia.Media.Brushes.Green;
        }
    }

    #endregion

    #region 水灌效果相关

    // 初始化动画（简化版）
    private void InitializeAnimations()
    {
        // TODO: 实现 Avalonia 动画
        // WPF 的 DoubleAnimation 在 Avalonia 中需要使用不同的方式
    }

    // 更新水灌效果
    private void UpdateWaterFill(double value, int waterFillnumber)
    {
        double maxValue;
        double percentage;
        double height;

        if (waterFillnumber == 0)
        {
            // 学习累计水灌
            maxValue = 520;
            percentage = Math.Min(value / maxValue, 1.0);

            if (_waterContainer != null && _waterFill != null)
            {
                height = _waterContainer.Bounds.Height * percentage;
                _waterFill.Height = height;

                if (IsHuanzhai && total2 >= total)
                {
                    total2 = total2 - total;
                    total = 0;
                    percentage = Math.Min(total / maxValue, 1.0);
                    height = _waterContainer.Bounds.Height * percentage;
                    _waterFill.Height = height;
                    UpdateWaterColor(total);
                    UpdateWaterColor2(total2);
                }
                else
                {
                    UpdateWaterColor(value);
                    UpdateWaterColor2(total2);
                }
            }
        }

        if (waterFillnumber == 2)
        {
            // 娱乐累计水灌
            maxValue = 520;
            percentage = Math.Min(value / maxValue, 1.0);

            if (_waterContainer2 != null && _waterFill2 != null)
            {
                height = _waterContainer2.Bounds.Height * percentage;
                _waterFill2.Height = height;

                if (IsHuanzhai && total2 >= total)
                {
                    total2 = total2 - total;
                    total = 0;
                    percentage = Math.Min(total2 / maxValue, 1.0);
                    height = _waterContainer2.Bounds.Height * percentage;
                    _waterFill2.Height = height;
                    UpdateWaterColor(total);
                    UpdateWaterColor2(total2);
                }
                else
                {
                    UpdateWaterColor2(total2);
                }
            }
        }

        if (waterFillnumber == 1)
        {
            // 总累计水灌
            maxValue = 63000;
            percentage = Math.Min(value / maxValue, 1.0);

            if (_waterContainer1 != null && _waterFill1 != null)
            {
                height = _waterContainer1.Bounds.Height * percentage;
                _waterFill1.Height = height;
                UpdateWaterColor2(total2);
            }
        }
    }

    // 更新水的颜色（学习累计）
    private void UpdateWaterColor(double value)
    {
        if (_waterFill == null) return;

        var topColor = value <= 130 ? Color.FromRgb(255, 100, 100) :
                       value <= 260 ? Color.FromRgb(100, 255, 100) :
                       value <= 390 ? Color.FromRgb(100, 100, 255) :
                       Color.FromRgb(200, 100, 255);

        var bottomColor = value <= 130 ? Color.FromRgb(255, 0, 0) :
                          value <= 260 ? Color.FromRgb(0, 180, 0) :
                          value <= 390 ? Color.FromRgb(0, 0, 255) :
                          Color.FromRgb(128, 0, 128);

        var brush = _waterFill.Background as Avalonia.Media.LinearGradientBrush;
        if (brush != null && brush.GradientStops.Count >= 2)
        {
            brush.GradientStops[0].Color = topColor;
            brush.GradientStops[1].Color = bottomColor;
        }
    }

    // 更新水的颜色（娱乐累计）
    private void UpdateWaterColor2(double value)
    {
        if (_waterFill2 == null) return;

        var topColor = value >= total ? Color.FromRgb(255, 100, 100) : Color.FromRgb(100, 255, 100);
        var bottomColor = value >= total ? Color.FromRgb(255, 0, 0) : Color.FromRgb(0, 180, 0);

        var brush = _waterFill2.Background as Avalonia.Media.LinearGradientBrush;
        if (brush != null && brush.GradientStops.Count >= 2)
        {
            brush.GradientStops[0].Color = topColor;
            brush.GradientStops[1].Color = bottomColor;
        }
    }



    #endregion

    #region Excel 文件操作

    // 读取 Excel 第一个单元格的数字（完全复用场景 1 的解析逻辑）
    private double ReadFirstCellNumber(string filePath)
    {
        IWorkbook workbook = null;
        FileStream? fileStream = null;

        try
        {
            
            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                return 0;
            }

            // 根据文件扩展名，使用不同的类打开工作簿
            fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            string extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".xls")
            {
                workbook = new HSSFWorkbook(fileStream);
            }
            else if (extension == ".xlsx")
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            else
            {
                return 0;
            }

            // 获取第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 获取第一行，如果第一行不存在则返回0
            IRow row = sheet.GetRow(0);
            if (row == null)
            {
                return 0;
            }

            // 获取第一个单元格
            ICell cell = row.GetCell(0);
            if (cell == null)
            {
                return 0;
            }

            // 根据单元格类型获取值
            double value = 0;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    value = cell.NumericCellValue;
                    break;
                case CellType.String:
                    // 尝试将字符串转换为数字
                    if (double.TryParse(cell.StringCellValue, out double num))
                    {
                        value = num;
                    }
                    else
                    {
                        throw new Exception($"第一个单元格不是有效的数字: {cell.StringCellValue}");
                    }
                    break;
                case CellType.Formula:
                    // 公式单元格，获取计算后的值
                    if (cell.CachedFormulaResultType == CellType.Numeric)
                    {
                        value = cell.NumericCellValue;
                    }
                    else if (cell.CachedFormulaResultType == CellType.String)
                    {
                        if (double.TryParse(cell.StringCellValue, out double num2))
                        {
                            value = num2;
                        }
                        else
                        {
                            throw new Exception($"第一个单元格公式结果不是有效的数字: {cell.StringCellValue}");
                        }
                    }
                    else
                    {
                        throw new Exception($"第一个单元格公式结果类型不支持: {cell.CachedFormulaResultType}");
                    }
                    break;
                case CellType.Blank:
                    return 0;
                default:
                    throw new Exception($"第一个单元格类型不支持: {cell.CellType}");
            }

            return value;
        }
        catch (Exception ex)
        {
            // 异常兜底：任何步骤失败均返回 0，避免 App 崩溃
            return 0;
        }
        finally
        {
            // 关键：无论是否报错，都强制释放资源
            workbook?.Close();
            workbook?.Dispose();
            fileStream?.Close();
            fileStream?.Dispose();
        }
    }

    // 保存数值到 Excel 第一个单元格
    private void SaveValueToFirstCell(string filePath, double valueToSave)
    {
        IWorkbook workbook = null;
        FileStream fileStream = null;
        MemoryStream? memoryStream = null;

        try
        {
            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                throw new Exception("文件不存在");
            }

            // 以读写模式打开文件，保持流始终打开
            fileStream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None
            );

            // 读取文件内容到内存（关键：避免.xlsx依赖外部流）
            byte[] fileContent;
            using (var tempMemoryStream = new MemoryStream())
            {
                fileStream.CopyTo(tempMemoryStream);
                fileContent = tempMemoryStream.ToArray();
            }

            // 创建内存流用于工作簿
            memoryStream = new MemoryStream(fileContent);

            // 根据文件扩展名，使用不同的类打开工作簿
            string extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".xls")
            {
                workbook = new HSSFWorkbook(memoryStream);
            }
            else if (extension == ".xlsx")
            {
                workbook = new XSSFWorkbook(memoryStream);
            }
            else
            {
                throw new Exception($"不支持的文件格式: {extension}");
            }

            // 获取或创建工作表、行、单元格
            ISheet sheet = workbook.GetSheetAt(0) ?? workbook.CreateSheet("Sheet1");
            IRow row = sheet.GetRow(0) ?? sheet.CreateRow(0);
            ICell cell = row.GetCell(0) ?? row.CreateCell(0);

            // 设置单元格值
            cell.SetCellValue(valueToSave);

            // 写入修改后的数据到临时内存流
            using (var outputMemoryStream = new MemoryStream())
            {
                workbook.Write(outputMemoryStream);
                byte[] outputData = outputMemoryStream.ToArray();

                // 将修改后的数据写回原文件
                fileStream.Seek(0, SeekOrigin.Begin);
                fileStream.SetLength(0); // 清空原有内容
                fileStream.Write(outputData, 0, outputData.Length);
                fileStream.Flush(); // 强制写入磁盘
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"修改Excel内容失败：{ex.Message}", ex);
        }
        finally
        {
            // 关键：无论是否报错，都强制释放资源
            // 注意：workbook 必须在 memoryStream 关闭前关闭
            workbook?.Close();
            workbook?.Dispose();
            memoryStream?.Close();
            memoryStream?.Dispose();
            fileStream?.Close();
            fileStream?.Dispose();
        }
    }

    #endregion
}

// 数据类
public class ChengxiaoShuxin
{
    public double Chengxiaoleiji { get; set; }
    public double Yuleleiji { get; set; }
    public double Duoyuleiji { get; set; }
}

// 文件处理类
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
            var chengxiaoshuxin = data as ChengxiaoShuxin;
            
            Debug.WriteLine($"检测到 URI 格式，准备保存: {filePath}");
            Debug.WriteLine($"chengxiaoshuxin 是否为 null: {chengxiaoshuxin == null}");
            if (chengxiaoshuxin != null)
            {
                Debug.WriteLine($"Chengxiaoleiji={chengxiaoshuxin.Chengxiaoleiji}, Yuleleiji={chengxiaoshuxin.Yuleleiji}, Duoyuleiji={chengxiaoshuxin.Duoyuleiji}");
            }
            
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
                    
                    var saveMethod = handlerType.GetMethod("SaveDataToFile", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    if (saveMethod != null)
                    {
                        Debug.WriteLine($"找到 SaveDataToFile 方法，参数: {string.Join(", ", saveMethod.GetParameters().Select(p => p.ParameterType.Name))}");
                        Debug.WriteLine("调用 AndroidDataFileHandler.SaveDataToFile");
                        saveMethod.Invoke(null, new object[] { chengxiaoshuxin, filePath });
                        Debug.WriteLine("保存完成");
                    }
                    else
                    {
                        Debug.WriteLine("未找到 SaveDataToFile 方法");
                        // 列出所有方法
                        Debug.WriteLine("可用方法:");
                        foreach (var method in handlerType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
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
            // 其他平台：使用文本格式
            Debug.WriteLine($"使用普通文件路径: {filePath}");
            using (var writer = new StreamWriter(filePath))
            {
                var chengxiaoshuxin = data as ChengxiaoShuxin;
                writer.WriteLine($"Chengxiaoleiji:{chengxiaoshuxin?.Chengxiaoleiji}");
                writer.WriteLine($"Yuleleiji:{chengxiaoshuxin?.Yuleleiji}");
                writer.WriteLine($"Duoyuleiji:{chengxiaoshuxin?.Duoyuleiji}");
            }
        }
    }

    public static T LoadDataFromFile<T>(string filePath) where T : new()
    {
        var result = new T();
        var data = result as ChengxiaoShuxin;

        if (data == null) return result;

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
                    var loadMethod = handlerType.GetMethod("LoadDataFromFile", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    if (loadMethod != null)
                    {
                        Debug.WriteLine("调用 AndroidDataFileHandler.LoadDataFromFile");
                        var loadedData = loadMethod.Invoke(null, new object[] { filePath });

                        // 手动复制属性（避免类型转换问题）
                        if (loadedData != null)
                        {
                            var loadedProps = loadedData.GetType().GetProperties();
                            foreach (var prop in loadedProps)
                            {
                                var targetProp = data.GetType().GetProperty(prop.Name);
                                if (targetProp != null && targetProp.CanWrite)
                                {
                                    var value = prop.GetValue(loadedData);
                                    targetProp.SetValue(data, value);
                                    Debug.WriteLine($"复制属性 {prop.Name} = {value}");
                                }
                            }
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
            // 其他平台：使用文本格式
            Debug.WriteLine($"使用普通文件路径: {filePath}");
            using (var reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split(':');
                    if (parts.Length == 2)
                    {
                        if (parts[0] == "Chengxiaoleiji" && double.TryParse(parts[1], out double chengxiao))
                            data.Chengxiaoleiji = chengxiao;
                        else if (parts[0] == "Yuleleiji" && double.TryParse(parts[1], out double yule))
                            data.Yuleleiji = yule;
                        else if (parts[0] == "Duoyuleiji" && double.TryParse(parts[1], out double duoyule))
                            data.Duoyuleiji = duoyule;
                    }
                }
            }
        }

        return result;
    }
}
