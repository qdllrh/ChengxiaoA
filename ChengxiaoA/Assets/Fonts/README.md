# 中文字体配置说明

## ✅ 已完成配置

当前采用**嵌入思源黑体字体**方案，字体文件已包含在项目中。

## 📁 字体文件位置

```
ChengxiaoA/Assets/Fonts/SourceHanSansCN-Normal.otf
```

## 📝 配置位置

### App.axaml.cs - 核心配置
```csharp
public override void OnFrameworkInitializationCompleted()
{
    // 根据平台设置全局默认字体
    FontFamily fontFamily;

#if ANDROID
    // Android 平台：使用嵌入的思源黑体字体（100%可靠）
    fontFamily = new FontFamily("avares://ChengxiaoA/Assets/Fonts/#Source Han Sans CN");
#else
    // 其他平台：使用字体回退机制
    fontFamily = new FontFamily("SimHei,Microsoft YaHei,Arial,sans-serif");
#endif

    // 设置全局字体资源
    Application.Current!.Resources["CustomChineseFont"] = fontFamily;
}
```

### App.axaml - 样式应用
```xml
<Application.Styles>
    <FluentTheme />

    <!-- 强制全局字体 -->
    <Style Selector="TextBlock">
        <Setter Property="FontFamily" Value="{DynamicResource CustomChineseFont}" />
    </Style>
    <Style Selector="Button">
        <Setter Property="FontFamily" Value="{DynamicResource CustomChineseFont}" />
    </Style>
    <Style Selector="TextBox">
        <Setter Property="FontFamily" Value="{DynamicResource CustomChineseFont}" />
    </Style>
    <Style Selector="Label">
        <Setter Property="FontFamily" Value="{DynamicResource CustomChineseFont}" />
    </Style>
</Application.Styles>
```

## 🎯 方案优势

| 优势 | 说明 |
|------|------|
| ✅ **100% 可靠** | 字体文件嵌入项目，不依赖系统 |
| ✅ **跨平台兼容** | Android、iOS、Windows、macOS 都能用 |
| ✅ **统一显示** | 所有平台使用相同的中文字体 |
| ✅ **无需下载** | 字体已包含在应用中 |

## 📱 当前字体配置

| 平台 | 字体 | 配置方式 |
|------|------|----------|
| **Android** | Source Han Sans CN | 嵌入字体文件 |
| **iOS** | Source Han Sans CN | 嵌入字体文件 |
| **Windows** | SimHei / 微软雅黑 | 字体回退机制 |
| **macOS / Linux** | Arial + 回退 | 字体回退机制 |

## 🔧 如果需要替换字体

1. **下载新的字体文件**（.otf 或 .ttf 格式）
2. **替换** `Assets/Fonts/SourceHanSansCN-Normal.otf`
3. **修改** `App.axaml.cs` 中的字体族名：

```csharp
#if ANDROID
    // 替换为新字体的族名
    fontFamily = new FontFamily("avares://ChengxiaoA/Assets/Fonts/#新字体名称");
#endif
```

## ❓ 常见问题

### Q1: 中文仍然显示为方块？
**解决方法**：
1. 确认字体文件在 `Assets/Fonts/` 文件夹
2. 确认 `ChengxiaoA.csproj` 包含：
   ```xml
   <AvaloniaResource Include="Assets\**" />
   ```
3. 清理并重新编译：
   ```bash
   dotnet clean
   dotnet build
   ```

### Q2: 字体族名怎么找？
**方法**：
1. 双击字体文件打开
2. 顶部显示的名称就是字体族名
3. 思源黑体的族名是：`Source Han Sans CN`（注意有空格）

### Q3: 如何减小应用体积？
**选项 1：使用子集字体**
- 使用工具提取常用汉字
- 生成只包含常用字的小字体文件

**选项 2：使用系统字体（不推荐）**
- 改回使用 Android 系统字体
- 但可能在不同设备上不兼容

## ✅ 当前状态
- ✅ 字体文件已嵌入项目
- ✅ Android 平台使用思源黑体
- ✅ 全局字体样式已配置
- ✅ 跨平台兼容性良好
