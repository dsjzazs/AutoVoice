using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;

namespace AutoVoice
{
    public static class VoiceDiagnostic
    {
        public static string GetVoiceInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== 系统语音包诊断信息 ===");
            sb.AppendLine();

            try
            {
                using var synthesizer = new SpeechSynthesizer();
                var voices = synthesizer.GetInstalledVoices();

                sb.AppendLine($"检测到 {voices.Count} 个语音包:");
                sb.AppendLine();

                foreach (var voice in voices)
                {
                    var info = voice.VoiceInfo;
                    sb.AppendLine($"名称: {info.Name}");
                    sb.AppendLine($"描述: {info.Description}");
                    sb.AppendLine($"文化: {info.Culture}");
                    sb.AppendLine($"性别: {info.Gender}");
                    sb.AppendLine($"年龄: {info.Age}");
                    sb.AppendLine($"ID: {info.Id}");
                    sb.AppendLine($"启用状态: {voice.Enabled}");
                    sb.AppendLine("---");
                }

                // 检查是否有英语语音包
                var englishVoices = voices.Where(v => 
                    v.VoiceInfo.Culture.Name.StartsWith("en-", StringComparison.OrdinalIgnoreCase)).ToList();
                
                sb.AppendLine();
                sb.AppendLine($"英语语音包数量: {englishVoices.Count}");
                foreach (var voice in englishVoices)
                {
                    sb.AppendLine($"- {voice.VoiceInfo.Name} ({voice.VoiceInfo.Culture.Name})");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"获取语音信息时出错: {ex.Message}");
            }

            return sb.ToString();
        }

        public static List<string> GetAvailableVoiceNames()
        {
            try
            {
                using var synthesizer = new SpeechSynthesizer();
                var voices = synthesizer.GetInstalledVoices();
                return voices.Select(v => v.VoiceInfo.Name).ToList();
            }
            catch
            {
                return new List<string>();
            }
        }
    }
} 