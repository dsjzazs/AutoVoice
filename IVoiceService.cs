using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AutoVoice
{
    /// <summary>
    /// 语音服务接口，支持多种 TTS 引擎
    /// </summary>
    public interface IVoiceService : IDisposable
    {
        /// <summary>
        /// 设置语速 (0.5 - 2.0)
        /// </summary>
        double Speed { get; set; }

        /// <summary>
        /// 设置音量 (0 - 100)
        /// </summary>
        int Volume { get; set; }

        /// <summary>
        /// 当前使用的语音模型
        /// </summary>
        string CurrentModel { get; set; }

        /// <summary>
        /// 检查语音引擎是否可用
        /// </summary>
        Task<(bool isAvailable, string message)> IsAvailableAsync();

        /// <summary>
        /// 获取可用的语音列表
        /// </summary>
        Task<List<string>> GetAvailableVoicesAsync();

        /// <summary>
        /// 异步播放文本
        /// </summary>
        Task<(bool success, string message)> SpeakAsync(string text);

        /// <summary>
        /// 停止当前播放
        /// </summary>
        void Stop();

        /// <summary>
        /// 获取引擎名称
        /// </summary>
        string EngineName { get; }
    }
}
