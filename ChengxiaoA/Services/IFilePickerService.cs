using System.Threading.Tasks;

namespace ChengxiaoA.Services;

/// <summary>
/// 文件选择器接口 - 用于跨平台文件选择
/// </summary>
public interface IFilePickerService
{
    /// <summary>
    /// 选择 Excel 文件
    /// </summary>
    /// <returns>文件路径，如果取消则返回空字符串</returns>
    Task<string> PickExcelFileAsync();

    /// <summary>
    /// 保存文件
    /// </summary>
    /// <param name="defaultFileName">默认文件名</param>
    /// <returns>文件路径，如果取消则返回空字符串</returns>
    Task<string> SaveFileAsync(string defaultFileName);
}
public interface IFileHandler
{
    // 核心修改：返回类型从 int 改为 double
    double ReadExcelFromUri(string filePath);
}