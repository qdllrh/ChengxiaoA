using System.Text.Json.Serialization;

namespace ChengxiaoA.Models;

/// <summary>
/// 目标数据类
/// </summary>
public class MubaioShuxin
{
    /// <summary>
    /// 目标名称
    /// </summary>
    [JsonPropertyName("Name")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// 目标达到时间（小时）
    /// </summary>
    [JsonPropertyName("Time")]
    public string Time { get; set; } = string.Empty;

    /// <summary>
    /// 当前成效
    /// </summary>
    [JsonPropertyName("Dangqianchengxiao")]
    public double Dangqianchengxiao { get; set; }

    /// <summary>
    /// 比例
    /// </summary>
    [JsonPropertyName("Bili")]
    public double Bili { get; set; }

    /// <summary>
    /// 目标成效
    /// </summary>
    [JsonPropertyName("Mubiaochengxiao")]
    public double Mubiaochengxiao { get; set; }
}
