using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace AutoVoice
{
    public static class VoiceDiagnostic
    {
        public static string GetVoiceInfo(string engineType = "Piper")
        {
            return GetVoiceInfoAsync(engineType).GetAwaiter().GetResult();
        }

        public static async Task<string> GetVoiceInfoAsync(string engineType = "Piper")
        {
            var sb = new StringBuilder();

            if (engineType == "Piper")
            {
                sb.AppendLine("=== Piper TTS 语音包诊断信息 ===");
                sb.AppendLine();

                try
                {
                    var piperService = new PiperVoiceService();
                    
                    // 检查 Piper 是否已安装
                    var isPiperInstalled = await piperService.IsPiperInstalledAsync();
                    sb.AppendLine($"Piper TTS 安装状态: {(isPiperInstalled.isInstalled ? "已安装 ✓" : "未安装 ✗")}");
                    sb.AppendLine();

                    if (!isPiperInstalled.isInstalled)
                    {
                        sb.AppendLine("请先安装 Piper TTS:");
                        sb.AppendLine("  pip install piper-tts");
                        sb.AppendLine();
                        piperService.Dispose();
                        return sb.ToString();
                    }

                    // 获取可用语音列表
                    var voices = await piperService.GetAvailableVoicesAsync();
                    sb.AppendLine($"可用语音模型数量: {voices.Count}");
                    sb.AppendLine();

                    sb.AppendLine("语音模型列表:");
                    foreach (var voice in voices)
                    {
                        bool isDownloaded = piperService.IsVoiceDownloaded(voice);
                        string status = isDownloaded ? "已下载 ✓" : "未下载";
                        sb.AppendLine($"- {voice} ({status})");
                    }

                    sb.AppendLine();
                    sb.AppendLine("当前配置:");
                    sb.AppendLine($"- 当前模型: {piperService.CurrentModel}");
                    sb.AppendLine($"- 语速: {piperService.Speed}x");
                    sb.AppendLine($"- 音量: {piperService.Volume}%");
                    
                    piperService.Dispose();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"获取语音信息时出错: {ex.Message}");
                    sb.AppendLine($"详细错误: {ex.StackTrace}");
                }
            }
            else // Windows TTS
            {
                sb.AppendLine("=== Windows TTS 语音包诊断信息 ===");
                sb.AppendLine();

                try
                {
                    var windowsService = new WindowsVoiceService();
                    
                    var (isAvailable, message) = await windowsService.IsAvailableAsync();
                    sb.AppendLine($"Windows TTS 状态: {(isAvailable ? "可用 ✓" : "不可用 ✗")}");
                    sb.AppendLine($"消息: {message}");
                    sb.AppendLine();

                    if (isAvailable)
                    {
                        var voices = await windowsService.GetAvailableVoicesAsync();
                        sb.AppendLine($"可用语音数量: {voices.Count}");
                        sb.AppendLine();

                        sb.AppendLine("语音列表:");
                        foreach (var voice in voices)
                        {
                            sb.AppendLine($"- {voice}");
                        }

                        sb.AppendLine();
                        sb.AppendLine("当前配置:");
                        sb.AppendLine($"- 当前语音: {windowsService.CurrentModel}");
                        sb.AppendLine($"- 语速: {windowsService.Speed}x");
                        sb.AppendLine($"- 音量: {windowsService.Volume}%");
                    }
                    
                    windowsService.Dispose();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"获取语音信息时出错: {ex.Message}");
                    sb.AppendLine($"详细错误: {ex.StackTrace}");
                }
            }

            return sb.ToString();
        }

        public static async Task<List<string>> GetAvailableVoiceNamesAsync()
        {
            try
            {
                var piperService = new PiperVoiceService();
                var voices = await piperService.GetAvailableVoicesAsync();
                piperService.Dispose();
                return voices;
            }
            catch
            {
                return new List<string> { "en_US-lessac-medium" };
            }
        }
    }
}