using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChengxiaoA.Services
{
    /// <summary>
    /// Excel 读取服务接口（跨平台通用）
    /// </summary>
    public interface IExcelReaderService
    {
        /// <summary>
        /// 读取 Excel 文件第一个单元格的数值（兼容普通路径/Android URI）
        /// </summary>
        /// <param name="filePath">文件路径或 Android URI</param>
        /// <returns>第一个单元格的数值，失败返回 0</returns>
        double  ReadFirstCellNumber(string filePath);

        /// <summary>
        /// 写入数值到 Excel 文件的第一个单元格（兼容普通路径/Android URI）
        /// </summary>
        /// <param name="filePath">文件路径或 Android URI</param>
        /// <param name="value">要写入的数值</param>
        /// <returns>是否写入成功</returns>
        bool WriteFirstCellNumber(string filePath, double value);
    }
   
}
