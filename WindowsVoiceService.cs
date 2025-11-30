using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Threading.Tasks;

namespace AutoVoice
{
    /// <summary>
    /// Windows 内置语音服务封装类 (System.Speech.Synthesis)
    /// </summary>
    public class WindowsVoiceService : IVoiceService
    {
        private SpeechSynthesizer? synthesizer;
        private string currentVoice = "";
        private double speed = 1.0;
        private int volume = 100;
        private bool disposed = false;

        public WindowsVoiceService()
        {
            try
            {
                synthesizer = new SpeechSynthesizer();
                synthesizer.SetOutputToDefaultAudioDevice();
                
                // 获取默认语音
                var voices = synthesizer.GetInstalledVoices();
                if (voices.Count > 0)
                {
                    currentVoice = voices[0].VoiceInfo.Name;
                    synthesizer.SelectVoice(currentVoice);
                }
            }
            catch (Exception)
            {
                // 初始化失败
            }
        }

        public string EngineName => "Windows TTS";

        public double Speed
        {
            get => speed;
            set
            {
                speed = Math.Max(0.5, Math.Min(2.0, value));
                if (synthesizer != null)
                {
                    // Windows Speech 使用 -10 到 10 的范围，0为正常
                    // 将 0.5-2.0 映射到 -10 到 10
                    int rate = (int)Math.Round((speed - 1.0) * 10);
                    rate = Math.Max(-10, Math.Min(10, rate));
                    synthesizer.Rate = rate;
                }
            }
        }

        public int Volume
        {
            get => volume;
            set
            {
                volume = Math.Max(0, Math.Min(100, value));
                if (synthesizer != null)
                {
                    synthesizer.Volume = volume;
                }
            }
        }

        public string CurrentModel
        {
            get => currentVoice;
            set
            {
                if (!string.IsNullOrEmpty(value) && synthesizer != null)
                {
                    try
                    {
                        synthesizer.SelectVoice(value);
                        currentVoice = value;
                    }
                    catch
                    {
                        // 选择失败，保持当前语音
                    }
                }
            }
        }

        public Task<(bool isAvailable, string message)> IsAvailableAsync()
        {
            try
            {
                if (synthesizer == null)
                {
                    return Task.FromResult((false, "Windows TTS 未初始化"));
                }

                var voices = synthesizer.GetInstalledVoices();
                if (voices.Count == 0)
                {
                    return Task.FromResult((false, "未找到已安装的语音包"));
                }

                return Task.FromResult((true, $"Windows TTS 已就绪，共 {voices.Count} 个语音"));
            }
            catch (Exception ex)
            {
                return Task.FromResult((false, $"Windows TTS 检查失败: {ex.Message}"));
            }
        }

        public Task<List<string>> GetAvailableVoicesAsync()
        {
            var voices = new List<string>();
            
            try
            {
                if (synthesizer != null)
                {
                    foreach (var voice in synthesizer.GetInstalledVoices())
                    {
                        if (voice.Enabled)
                        {
                            voices.Add(voice.VoiceInfo.Name);
                        }
                    }
                }
            }
            catch
            {
                // 如果获取失败，返回空列表
            }

            return Task.FromResult(voices);
        }

        public Task<(bool success, string message)> SpeakAsync(string text)
        {
            if (synthesizer == null)
            {
                return Task.FromResult((false, "语音合成器未初始化"));
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                return Task.FromResult((false, "文本为空"));
            }

            try
            {
                // 停止当前播放
                Stop();

                // 使用异步方式播放
                return Task.Run(() =>
                {
                    try
                    {
                        synthesizer.Speak(text);
                        return (true, "播放成功");
                    }
                    catch (Exception ex)
                    {
                        return (false, $"播放失败: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                return Task.FromResult((false, $"播放出错: {ex.Message}"));
            }
        }

        public void Stop()
        {
            try
            {
                synthesizer?.SpeakAsyncCancelAll();
            }
            catch
            {
                // 忽略停止失败
            }
        }

        public void Dispose()
        {
            if (!disposed)
            {
                Stop();
                synthesizer?.Dispose();
                synthesizer = null;
                disposed = true;
            }
        }
    }
}
