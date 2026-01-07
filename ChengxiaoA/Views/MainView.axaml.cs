using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform;
using ChengxiaoA.Services;
using System;
using System.IO;
using System.Diagnostics;
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
    public static double Duoyuleji;
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
        var Duoyujileibutton = this.FindControl<Button>("Duoyujileibutton");

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

                if (total > total2 && Duoyujileibutton != null)
                {
                    Duoyujileibutton.IsEnabled = true;
                }
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
        var Duoyujileibutton = this.FindControl<Button>("Duoyujileibutton");

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

                if (total > total2 && Duoyujileibutton != null)
                {
                    Duoyujileibutton.IsEnabled = true;
                }
            }
            else
            {
                Debug.WriteLine("请先点击「娱乐」获取有效结果！");
            }
        }
    }

    #endregion

    #region 多余积累事件

    private void Duoyujilei_Click(object? sender, RoutedEventArgs e)
    {
        var txtDuoyuleiji = this.FindControl<TextBox>("txtDuoyuleiji");
        var txtleijiYule = this.FindControl<TextBox>("txtleijiYule");
        var Duoyujileibutton = this.FindControl<Button>("Duoyujileibutton");

        if (total > total2)
        {
            Duoyuleji = Duoyuleji + (total - total2);

            if (txtDuoyuleiji != null)
                txtDuoyuleiji.Text = Duoyuleji.ToString("F2"); // 显示累计结果

            if (Duoyujileibutton != null)
                Duoyujileibutton.IsEnabled = false;

            total2 = total;

            if (txtleijiYule != null)
                txtleijiYule.Text = total2.ToString("F2"); // 显示累计结果

            UpdateWaterFill(total2, 2);
            UpdateWaterColor2(total2);
        }
        else
        {
            Debug.WriteLine("学习累计不足，先累计学习");
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
#if ANDROID
        // Android 平台：使用存储访问框架
        return await FilePickerService.PickExcelFileAsync();
#else
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
#endif
    }

    private void ZongjileiJisuan_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");
            var txtDuoyuleiji = this.FindControl<TextBox>("txtDuoyuleiji");

            if (string.IsNullOrEmpty(_selectedExcelPath))
            {
                Debug.WriteLine("请先选择Excel文件！");
                return;
            }

            double.TryParse(txtXuexizongjilei?.Text, out double valueToSave);
            valueToSave = valueToSave + Duoyuleji;

            SaveValueToFirstCell(_selectedExcelPath, valueToSave);

            if (txtXuexizongjilei != null)
                txtXuexizongjilei.Text = $"{valueToSave}";

            Duoyuleji = 0;
            if (txtDuoyuleiji != null)
                txtDuoyuleiji.Text = Duoyuleji.ToString("F2");

            UpdateWaterFill(Duoyuleji, 1);
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
                Yuleleiji = total2,
                Duoyuleiji = Duoyuleji
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
                Duoyuleji = gongyongShuxing1.Duoyuleiji;
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
        var txtDuoyuleiji = this.FindControl<TextBox>("txtDuoyuleiji");
        var txtXuexizongjilei = this.FindControl<TextBox>("txtXuexizongjilei");

        if (txtTotal != null)
            txtTotal.Text = total.ToString("F2");

        UpdateWaterFill(total, 0);

        if (txtleijiYule != null)
            txtleijiYule.Text = total2.ToString("F2");

        UpdateWaterFill(total2, 2);

        if (txtDuoyuleiji != null)
            txtDuoyuleiji.Text = Duoyuleji.ToString("F2");

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
            // 多余累计水灌
            maxValue = 63000;
            percentage = Math.Min(value / maxValue, 1.0);

            if (waterContainer1 != null && waterFill1 != null)
            {
                height = waterContainer1.Bounds.Height * percentage;
                waterFill1.Height = height;
                UpdateWaterColor1(Duoyuleji);
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

    // 更新水的颜色（多余累计）
    private void UpdateWaterColor1(double value)
    {
        var waterFill1 = this.FindControl<Border>("waterFill1");
        if (waterFill1 == null) return;

        var topColor = Color.FromRgb(100, 255, 100);
        var bottomColor = Color.FromRgb(0, 180, 0);

        var brush = waterFill1.Background as Avalonia.Media.LinearGradientBrush;
        if (brush != null && brush.GradientStops.Count >= 2)
        {
            brush.GradientStops[0].Color = topColor;
            brush.GradientStops[1].Color = bottomColor;
        }
    }

    #endregion

    #region Excel 文件操作

    // 读取 Excel 第一个单元格的数字
    private double ReadFirstCellNumber(string filePath)
    {
        IWorkbook workbook = null; // 先定义为null，方便后续释放

        try
        {
            // 根据文件扩展名，使用不同的类打开工作簿
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (System.IO.Path.GetExtension(filePath).ToLower() == ".xls")
                {
                    workbook = new HSSFWorkbook(file);
                }
                else
                {
                    workbook = new XSSFWorkbook(file);
                }
            }

            // 获取第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 获取第一行，如果第一行不存在则返回0
            IRow row = sheet.GetRow(0);
            if (row == null)
            {
                Debug.WriteLine("第一行不存在");
                return 0;
            }

            // 获取第一个单元格
            ICell cell = row.GetCell(0);
            if (cell == null)
            {
                Debug.WriteLine("第一个单元格不存在");
                return 0;
            }

            // 根据单元格类型获取值
            double value = 0;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    value = cell.NumericCellValue;
                    Debug.WriteLine($"读取到数值: {value}");
                    break;
                case CellType.String:
                    // 尝试将字符串转换为数字
                    if (double.TryParse(cell.StringCellValue, out double num))
                    {
                        value = num;
                        Debug.WriteLine($"读取到字符串数值: {value}");
                    }
                    else
                    {
                        throw new Exception("第一个单元格不是有效的数字。");
                    }
                    break;
                case CellType.Formula:
                    // 公式单元格，我们需要计算值，但这里简单处理：如果是数字公式，则获取计算后的值
                    if (cell.CachedFormulaResultType == CellType.Numeric)
                    {
                        value = cell.NumericCellValue;
                        Debug.WriteLine($"读取到公式数值: {value}");
                    }
                    else if (cell.CachedFormulaResultType == CellType.String)
                    {
                        if (double.TryParse(cell.StringCellValue, out double num2))
                        {
                            value = num2;
                            Debug.WriteLine($"读取到公式字符串数值: {value}");
                        }
                        else
                        {
                            throw new Exception("第一个单元格公式结果不是有效的数字。");
                        }
                    }
                    else
                    {
                        throw new Exception("第一个单元格公式结果不是数字。");
                    }
                    break;
                default:
                    throw new Exception("第一个单元格不是数字类型。");
            }

            return value;
        }
        finally // 关键：无论是否报错，都强制释放workbook资源
        {
            workbook?.Close(); // 关闭工作簿
            workbook?.Dispose(); // 彻底释放资源（NPOI推荐的释放方式）
        }
    }

    // 保存数值到 Excel 第一个单元格
    private void SaveValueToFirstCell(string filePath, double valueToSave)
    {
        IWorkbook workbook = null;
        FileStream fileStream = null;

        try
        {
            // 以读写模式打开文件，保持流始终打开
            fileStream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None
            );

            // 读取文件内容到内存（关键：避免.xlsx依赖外部流）
            byte[] fileContent;
            using (var memoryStream = new MemoryStream())
            {
                fileStream.CopyTo(memoryStream);
                fileContent = memoryStream.ToArray();
            }

            // 从内存中创建工作簿（脱离对原文件流的依赖）
            using (var memoryStream = new MemoryStream(fileContent))
            {
                if (System.IO.Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new HSSFWorkbook(memoryStream);
                }
                else
                {
                    workbook = new XSSFWorkbook(memoryStream);
                }
            }

            // 获取或创建工作表、行、单元格
            ISheet sheet = workbook.GetSheetAt(0) ?? workbook.CreateSheet("Sheet1");
            IRow row = sheet.GetRow(0) ?? sheet.CreateRow(0);
            ICell cell = row.GetCell(0) ?? row.CreateCell(0);
            cell.SetCellValue(valueToSave);

            // 写入修改后的数据
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.SetLength(0); // 清空原有内容
            workbook.Write(fileStream);
            fileStream.Flush(); // 强制写入磁盘

            Debug.WriteLine($"成功保存数值 {valueToSave} 到 Excel 文件");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"修改Excel内容失败：{ex.Message}");
            throw new Exception($"修改Excel内容失败：{ex.Message}", ex);
        }
        finally
        {
            // 释放资源
            workbook?.Close();
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
        // TODO: 实现数据保存功能
        using (var writer = new StreamWriter(filePath))
        {
            writer.WriteLine($"Chengxiaoleiji:{(data as ChengxiaoShuxin).Chengxiaoleiji}");
            writer.WriteLine($"Yuleleiji:{(data as ChengxiaoShuxin).Yuleleiji}");
            writer.WriteLine($"Duoyuleiji:{(data as ChengxiaoShuxin).Duoyuleiji}");
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
                    else if (parts[0] == "Duoyuleiji" && double.TryParse(parts[1], out double duoyule))
                        data.Duoyuleiji = duoyule;
                }
            }
        }

        return result;
    }
}
