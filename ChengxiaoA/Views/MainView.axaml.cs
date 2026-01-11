using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform;
using ChengxiaoA.Services;
using System;
using System.IO;
using System.Diagnostics;
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

    public MainView()
    {
        InitializeComponent();

        // 注册 Loaded 事件
        this.Loaded += MainWindow_Loaded;

        // 初始化水灌动画
        InitializeAnimations();

        // 调试：打印字体信息
        Debug.WriteLine("========== 字体调试信息 ==========");
        Debug.WriteLine($"UserControl FontFamily: {this.FontFamily}");

        var titleBlock = this.FindControl<TextBlock>("titleTextBlock");
        if (titleBlock != null)
        {
            Debug.WriteLine($"标题 '成效计算器' FontFamily: {titleBlock.FontFamily}");
        }

        var calculateBtn = this.FindControl<Button>("btnCalculate");
        if (calculateBtn != null)
        {
            Debug.WriteLine($"按钮 '计算成效' FontFamily: {calculateBtn.FontFamily}");
        }
        Debug.WriteLine("=================================");
    }

    #region 窗口加载和关闭事件

    // 窗口加载事件 - 读取 Excel 文件
    private async void MainWindow_Loaded(object? sender, RoutedEventArgs e)
    {
        Debug.WriteLine("===== MainWindow_Loaded 被调用 =====");

        try
        {
            Debug.WriteLine("准备调用 PickExcelFileAsync...");
            // 调用平台特定的文件选择器
            string excelFilePath = await PickExcelFileAsync();

            Debug.WriteLine($"默认 Excel 路径: {excelFilePath}");
            //Debug.WriteLine($"文件存在: {File.Exists(excelFilePath)}");
            Debug.WriteLine($"文件存在: {excelFilePath}");
            //if (File.Exists(excelFilePath))
            if (!string.IsNullOrEmpty(excelFilePath))
                {
                var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
                if (txtXuexizongjilei != null)
                {
                    txtXuexizongjilei.Text = "正在读取文件...";

                    // 读取第一个单元格的数字
                    double firstCellNumber = ReadFirstCellNumber(excelFilePath);

                    // 显示结果
                    txtXuexizongjilei.Text = $"{firstCellNumber}";
                    Zongjilei1 = firstCellNumber;
                    _selectedExcelPath1 = excelFilePath;

                    Debug.WriteLine($"成功读取 Excel 文件: {excelFilePath}, 值: {firstCellNumber}");
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
                Debug.WriteLine($"默认 Excel 文件不存在: {excelFilePath}");
            }
        }
        catch (Exception ex)
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
            if (txtXuexizongjilei != null)
            {
                txtXuexizongjilei.Text = "读取失败，请检查文件格式";
            }
            Debug.WriteLine($"读取失败：{ex.Message}");
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
                Debug.WriteLine("请先选择Excel文件！");
                return;
            }

            SaveValueToFirstCell(_selectedExcelPath1, Zongjilei1);
            Debug.WriteLine($"总累计 {Zongjilei1} 已成功保存到A1单元格！");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"保存失败：{ex.Message}");
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
                    Debug.WriteLine("Y不能为0，请输入有效数值！");
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
            else
            {
                Debug.WriteLine("请确保X、Y、Z输入的是有效数值！");
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
            else
            {
                Debug.WriteLine("请先点击「计算成效」获取有效结果！");
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
            else
            {
                Debug.WriteLine("请确保娱乐时间输入的是有效数值！");
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
            else
            {
                Debug.WriteLine("请先点击「娱乐」获取有效结果！");
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
            Debug.WriteLine($"读取失败：{ex.Message}");
        }
    }

    // 平台特定的文件选择器
    private async Task<string> PickExcelFileAsync()
    {
        Debug.WriteLine("========== PickExcelFileAsync (MainView) 开始 ==========");

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

        Debug.WriteLine($"平台检测结果: {(isAndroid ? "Android" : "非Android")}");

        if (isAndroid)
        {
            Debug.WriteLine("✅ 检测到 Android 平台，调用 FilePickerService.PickExcelFileAsync()");
            // Android 平台：使用存储访问框架
            return await FilePickerService.PickExcelFileAsync();
        }
        else
        {
            Debug.WriteLine("❌ 非 Android 平台，使用默认文件选择器");
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
                Debug.WriteLine("请先选择Excel文件！");
                return;
            }

            double.TryParse(txtXuexizongjilei?.Text, out double valueToSave);

            SaveValueToFirstCell(_selectedExcelPath, valueToSave);

            if (txtXuexizongjilei != null)
                txtXuexizongjilei.Text = $"{valueToSave}";

            Debug.WriteLine($"数值 {valueToSave} 已成功保存到A1单元格！");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"保存失败：{ex.Message}");
        }
    }

    #endregion

    #region 保存和加载事件

    private async void Baocun_Click(object? sender, RoutedEventArgs e)
    {
        // TODO: 在 Android 上实现文件保存对话框
        // 暂时保存到默认路径
        string filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "chengxiao_data.dat"
        );

        try
        {
            // 创建数据对象
            var chengxiaoshuxin = new ChengxiaoShuxin
            {
                Chengxiaoleiji = total,
                Yuleleiji = total2
            };

            // 保存到文件
            DataFileHandler.SaveDataToFile(chengxiaoshuxin, filePath);
            Debug.WriteLine($"文件已保存到: {filePath}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"保存失败：{ex.Message}");
        }
    }

    private void Jiazai_Click(object? sender, RoutedEventArgs e)
    {
        // TODO: 在 Android 上实现文件选择对话框
        // 暂时从默认路径加载
        string filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "chengxiao_data.dat"
        );

        try
        {
            if (File.Exists(filePath))
            {
                var gongyongShuxing1 = DataFileHandler.LoadDataFromFile<ChengxiaoShuxin>(filePath);
                total = gongyongShuxing1.Chengxiaoleiji;
                total2 = gongyongShuxing1.Yuleleiji;
                Jiazaixiaoguo();
            }
            else
            {
                Debug.WriteLine("文件不存在");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"加载失败：{ex.Message}");
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
        // TODO: 打开目标设置窗口
        Debug.WriteLine("打开目标设置窗口");
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
        Debug.WriteLine("水灌动画初始化（待实现）");
    }

    // 更新水灌效果
    private void UpdateWaterFill(double value, int waterFillnumber)
    {
        var waterContainer = this.FindControl<Grid>("waterContainer");
        var waterFill = this.FindControl<Border>("waterFill");
        var waterContainer1 = this.FindControl<Grid>("waterContainer1");
        var waterFill1 = this.FindControl<Border>("waterFill1");
        var waterContainer2 = this.FindControl<Grid>("waterContainer2");
        var waterFill2 = this.FindControl<Border>("waterFill2");

        double maxValue;
        double percentage;
        double height;

        if (waterFillnumber == 0)
        {
            // 学习累计水灌
            maxValue = 520;
            percentage = Math.Min(value / maxValue, 1.0);

            if (waterContainer != null && waterFill != null)
            {
                height = waterContainer.Bounds.Height * percentage;
                waterFill.Height = height;

                if (IsHuanzhai && total2 >= total)
                {
                    total2 = total2 - total;
                    total = 0;
                    percentage = Math.Min(total / maxValue, 1.0);
                    height = waterContainer.Bounds.Height * percentage;
                    waterFill.Height = height;
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

            if (waterContainer2 != null && waterFill2 != null)
            {
                height = waterContainer2.Bounds.Height * percentage;
                waterFill2.Height = height;

                if (IsHuanzhai && total2 >= total)
                {
                    total2 = total2 - total;
                    total = 0;
                    percentage = Math.Min(total2 / maxValue, 1.0);
                    height = waterContainer2.Bounds.Height * percentage;
                    waterFill2.Height = height;
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

            if (waterContainer1 != null && waterFill1 != null)
            {
                height = waterContainer1.Bounds.Height * percentage;
                waterFill1.Height = height;
                UpdateWaterColor2(total2);
            }
        }
    }

    // 更新水的颜色（学习累计）
    private void UpdateWaterColor(double value)
    {
        var waterFill = this.FindControl<Border>("waterFill");
        if (waterFill == null) return;

        var topColor = value <= 130 ? Color.FromRgb(255, 100, 100) :
                       value <= 260 ? Color.FromRgb(100, 255, 100) :
                       value <= 390 ? Color.FromRgb(100, 100, 255) :
                       Color.FromRgb(200, 100, 255);

        var bottomColor = value <= 130 ? Color.FromRgb(255, 0, 0) :
                          value <= 260 ? Color.FromRgb(0, 180, 0) :
                          value <= 390 ? Color.FromRgb(0, 0, 255) :
                          Color.FromRgb(128, 0, 128);

        var brush = waterFill.Background as Avalonia.Media.LinearGradientBrush;
        if (brush != null && brush.GradientStops.Count >= 2)
        {
            brush.GradientStops[0].Color = topColor;
            brush.GradientStops[1].Color = bottomColor;
        }
    }

    // 更新水的颜色（娱乐累计）
    private void UpdateWaterColor2(double value)
    {
        var waterFill2 = this.FindControl<Border>("waterFill2");
        if (waterFill2 == null) return;

        var topColor = value >= total ? Color.FromRgb(255, 100, 100) : Color.FromRgb(100, 255, 100);
        var bottomColor = value >= total ? Color.FromRgb(255, 0, 0) : Color.FromRgb(0, 180, 0);

        var brush = waterFill2.Background as Avalonia.Media.LinearGradientBrush;
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
        Debug.WriteLine($"========== ReadFirstCellNumber 开始 ==========");
        Debug.WriteLine($"文件路径: {filePath}");

        IWorkbook workbook = null;
        FileStream? fileStream = null;

        try
        {
            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                Debug.WriteLine("❌ 文件不存在");
                return 0;
            }

            // 根据文件扩展名，使用不同的类打开工作簿
            fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            string extension = Path.GetExtension(filePath).ToLower();
            Debug.WriteLine($"文件扩展名: {extension}");

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
                Debug.WriteLine($"❌ 不支持的文件格式: {extension}");
                return 0;
            }

            Debug.WriteLine("✅ 工作簿创建成功");

            // 获取第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);
            Debug.WriteLine($"✅ 获取工作表: {sheet.SheetName}");

            // 获取第一行，如果第一行不存在则返回0
            IRow row = sheet.GetRow(0);
            if (row == null)
            {
                Debug.WriteLine("⚠️ 第一行不存在");
                return 0;
            }

            // 获取第一个单元格
            ICell cell = row.GetCell(0);
            if (cell == null)
            {
                Debug.WriteLine("⚠️ 第一个单元格不存在");
                return 0;
            }

            Debug.WriteLine($"单元格类型: {cell.CellType}");

            // 根据单元格类型获取值
            double value = 0;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    value = cell.NumericCellValue;
                    Debug.WriteLine($"✅ 读取到数值: {value}");
                    break;
                case CellType.String:
                    // 尝试将字符串转换为数字
                    if (double.TryParse(cell.StringCellValue, out double num))
                    {
                        value = num;
                        Debug.WriteLine($"✅ 读取到字符串数值: {value}");
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
                        Debug.WriteLine($"✅ 读取到公式数值: {value}");
                    }
                    else if (cell.CachedFormulaResultType == CellType.String)
                    {
                        if (double.TryParse(cell.StringCellValue, out double num2))
                        {
                            value = num2;
                            Debug.WriteLine($"✅ 读取到公式字符串数值: {value}");
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
                    Debug.WriteLine("⚠️ 第一个单元格为空");
                    return 0;
                default:
                    throw new Exception($"第一个单元格类型不支持: {cell.CellType}");
            }

            return value;
        }
        catch (Exception ex)
        {
            // 异常兜底：任何步骤失败均返回 0，避免 App 崩溃
            Debug.WriteLine($"❌ 读取 Excel 文件失败: {ex.Message}");
            Debug.WriteLine($"   异常类型: {ex.GetType().Name}");
            Debug.WriteLine($"   堆栈跟踪: {ex.StackTrace}");
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
        Debug.WriteLine($"========== SaveValueToFirstCell 开始 ==========");
        Debug.WriteLine($"文件路径: {filePath}");
        Debug.WriteLine($"要保存的值: {valueToSave}");

        IWorkbook workbook = null;
        FileStream fileStream = null;
        MemoryStream? memoryStream = null;

        try
        {
            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                Debug.WriteLine("❌ 文件不存在");
                throw new Exception("文件不存在");
            }

            Debug.WriteLine($"文件大小: {new FileInfo(filePath).Length} 字节");

            // 以读写模式打开文件，保持流始终打开
            fileStream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None
            );

            Debug.WriteLine("✅ 文件流打开成功");

            // 读取文件内容到内存（关键：避免.xlsx依赖外部流）
            byte[] fileContent;
            using (var tempMemoryStream = new MemoryStream())
            {
                fileStream.CopyTo(tempMemoryStream);
                fileContent = tempMemoryStream.ToArray();
                Debug.WriteLine($"✅ 读取文件内容到内存: {fileContent.Length} 字节");
            }

            // 创建内存流用于工作簿
            memoryStream = new MemoryStream(fileContent);

            // 根据文件扩展名，使用不同的类打开工作簿
            string extension = Path.GetExtension(filePath).ToLower();
            Debug.WriteLine($"文件扩展名: {extension}");

            if (extension == ".xls")
            {
                workbook = new HSSFWorkbook(memoryStream);
                Debug.WriteLine("✅ HSSFWorkbook 创建成功");
            }
            else if (extension == ".xlsx")
            {
                workbook = new XSSFWorkbook(memoryStream);
                Debug.WriteLine("✅ XSSFWorkbook 创建成功");
            }
            else
            {
                Debug.WriteLine($"❌ 不支持的文件格式: {extension}");
                throw new Exception($"不支持的文件格式: {extension}");
            }

            // 获取或创建工作表、行、单元格
            ISheet sheet = workbook.GetSheetAt(0) ?? workbook.CreateSheet("Sheet1");
            IRow row = sheet.GetRow(0) ?? sheet.CreateRow(0);
            ICell cell = row.GetCell(0) ?? row.CreateCell(0);

            // 设置单元格值
            cell.SetCellValue(valueToSave);
            Debug.WriteLine($"✅ 设置单元格值: {valueToSave}");

            // 写入修改后的数据到临时内存流
            using (var outputMemoryStream = new MemoryStream())
            {
                workbook.Write(outputMemoryStream);
                byte[] outputData = outputMemoryStream.ToArray();
                Debug.WriteLine($"✅ 工作簿写入内存: {outputData.Length} 字节");

                // 将修改后的数据写回原文件
                fileStream.Seek(0, SeekOrigin.Begin);
                fileStream.SetLength(0); // 清空原有内容
                fileStream.Write(outputData, 0, outputData.Length);
                fileStream.Flush(); // 强制写入磁盘
                Debug.WriteLine($"✅ 数据写入文件成功");
            }

            Debug.WriteLine($"✅ 成功保存数值 {valueToSave} 到 Excel 文件");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ 修改Excel内容失败：{ex.Message}");
            Debug.WriteLine($"   异常类型: {ex.GetType().Name}");
            Debug.WriteLine($"   堆栈跟踪: {ex.StackTrace}");
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
            Debug.WriteLine("========== SaveValueToFirstCell 结束 ==========");
        }
    }

    #endregion
}

// 数据类
public class ChengxiaoShuxin
{
    public double Chengxiaoleiji { get; set; }
    public double Yuleleiji { get; set; }
}

// 文件处理类
public static class DataFileHandler
{
    public static void SaveDataToFile<T>(T data, string filePath)
    {
        // TODO: 实现数据保存功能
        using (var writer = new StreamWriter(filePath))
        {
            writer.WriteLine($"Chengxiaoleiji:{(data as ChengxiaoShuxin).Chengxiaoleiji}");
            writer.WriteLine($"Yuleleiji:{(data as ChengxiaoShuxin).Yuleleiji}");
        }
    }

    public static T LoadDataFromFile<T>(string filePath) where T : new()
    {
        var result = new T();
        var data = result as ChengxiaoShuxin;

        if (data == null) return result;

        // TODO: 实现数据加载功能
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
                }
            }
        }

        return result;
    }
}
