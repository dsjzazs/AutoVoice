using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoVoice
{
    /// <summary>
    /// Piper TTS 语音服务封装类
    /// </summary>
    public class PiperVoiceService : IVoiceService
    {
        private string pythonPath = "python";
        private string dataDir;
        private string currentModel = "en_US-lessac-medium";
        private int volume = 100;
        private double speed = 1.0;
        private bool useNpu = false; // 启用 NPU 加速
        private bool disposed = false;

        public PiperVoiceService()
        {
            // 设置数据目录为应用程序目录下的 piper_data
            dataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "piper_data");
            Directory.CreateDirectory(dataDir);

            // 检测可用的 Python 路径
            pythonPath = DetectPythonPath();
        }

        public string EngineName => "Piper TTS";

        /// <summary>
        /// 检测可用的 Python 路径
        /// </summary>
        private string DetectPythonPath()
        {
            string[] candidates = { "python", "python3", "py" };
            foreach (var candidate in candidates)
            {
                try
                {
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = candidate,
                            Arguments = "--version",
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    process.Start();
                    process.WaitForExit();
                    if (process.ExitCode == 0)
                    {
                        return candidate;
                    }
                }
                catch
                {
                    // 忽略异常，继续尝试下一个
                }
            }
            // 如果都没有找到，返回默认值
            return "python";
        }

        /// <summary>
        /// 设置语速 (0.5 - 2.0)
        /// </summary>
        public double Speed
        {
            get => speed;
            set => speed = Math.Max(0.5, Math.Min(2.0, value));
        }

        /// <summary>
        /// 设置音量 (0 - 100)
        /// </summary>
        public int Volume
        {
            get => volume;
            set => volume = Math.Max(0, Math.Min(100, value));
        }

        /// <summary>
        /// 当前使用的语音模型
        /// </summary>
        public string CurrentModel
        {
            get => currentModel;
            set => currentModel = value ?? "en_US-lessac-medium";
        }

        /// <summary>
        /// 是否使用 NPU 加速
        /// </summary>
        public bool UseNpu
        {
            get => useNpu;
            set => useNpu = value;
        }

        /// <summary>
        /// 检查 Piper 是否可用
        /// </summary>
        public async Task<(bool isAvailable, string message)> IsAvailableAsync()
        {
            return await IsPiperInstalledAsync();
        }

        /// <summary>
        /// 检查 Piper 是否已安装
        /// </summary>
        public async Task<(bool isInstalled, string message)> IsPiperInstalledAsync()
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = pythonPath,
                        Arguments = "-c \"import piper\"",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode == 0)
                {
                    return (true, "Piper 已安装");
                }
                else
                {
                    string failureMessage = $"Piper 检查失败: ExitCode={process.ExitCode}, Error={error.Trim()}, Output={output.Trim()}";
                    return (false, failureMessage);
                }
            }
            catch (Exception ex)
            {
                string exceptionMessage = $"Piper 检查异常: {ex.Message}";
                return (false, exceptionMessage);
            }
        }

        /// <summary>
        /// 安装 Piper TTS
        /// </summary>
        public async Task<(bool success, string message)> InstallPiperAsync()
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = pythonPath,
                        Arguments = "-m pip install piper-tts",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode == 0)
                {
                    return (true, "Piper TTS 安装成功");
                }
                else
                {
                    return (false, $"安装失败 (ExitCode={process.ExitCode}): {error}\n输出: {output}");
                }
            }
            catch (Exception ex)
            {
                return (false, $"安装出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取可用的语音列表
        /// </summary>
        public async Task<List<string>> GetAvailableVoicesAsync()
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = pythonPath,
                        Arguments = "-m piper.download_voices",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync();

                // 解析输出获取语音列表
                var voices = new List<string>();
                foreach (var line in output.Split('\n'))
                {
                    if (line.Contains("en_") || line.Contains("zh_"))
                    {
                        var parts = line.Trim().Split(' ');
                        if (parts.Length > 0 && !string.IsNullOrWhiteSpace(parts[0]))
                        {
                            voices.Add(parts[0]);
                        }
                    }
                }

                return voices.Any() ? voices : new List<string> { "en_US-lessac-medium", "en_US-amy-medium", "en_GB-alan-medium" };
            }
            catch
            {
                // 返回默认列表
                return new List<string> { "en_US-lessac-medium", "en_US-amy-medium", "en_GB-alan-medium" };
            }
        }

        /// <summary>
        /// 下载指定的语音模型
        /// </summary>
        public async Task<(bool success, string message)> DownloadVoiceAsync(string voiceName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = pythonPath,
                        Arguments = $"-m piper.download_voices {voiceName} --data-dir \"{dataDir}\"",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode == 0 || output.Contains("already exists"))
                {
                    return (true, $"语音模型 {voiceName} 下载完成");
                }
                else
                {
                    return (false, $"下载失败: {error}");
                }
            }
            catch (Exception ex)
            {
                return (false, $"下载出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查语音模型是否已下载
        /// </summary>
        public bool IsVoiceDownloaded(string voiceName)
        {
            string modelPath = Path.Combine(dataDir, voiceName);
            return Directory.Exists(modelPath) && Directory.GetFiles(modelPath, "*.onnx").Any();
        }

        /// <summary>
        /// 异步播放文本
        /// </summary>
        public async Task<(bool success, string message)> SpeakAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return (false, "文本为空");
            }

            try
            {
                // 转义文本中的特殊字符
                string escapedText = text.Replace("\"", "\\\"");

                // 计算音量参数 (Piper 使用 0.0 - 1.0 的范围)
                double volumeMultiplier = volume / 100.0;

                // 生成临时文件
                var tempWavFile = Path.Combine(Path.GetTempPath(), $"piper_{Guid.NewGuid()}.wav");

                // 构建命令行参数 - 输出到临时文件
                string npuFlag = useNpu ? "--use-npu" : "";
                string arguments = $"-m piper -m {currentModel} --data-dir \"{dataDir}\" --volume {volumeMultiplier:F2} {npuFlag} -f \"{tempWavFile}\" -- \"{escapedText}\"";

                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = pythonPath,
                        Arguments = arguments,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                // 启动进程生成音频
                process.Start();
                
                // 异步等待生成完成
                var processTask = process.WaitForExitAsync();
                var errorTask = process.StandardError.ReadToEndAsync();

                // 等待进程结束
                await processTask;
                string error = await errorTask;

                if (process.ExitCode != 0)
                {
                    try { if (File.Exists(tempWavFile)) File.Delete(tempWavFile); } catch { }
                    return (false, $"TTS 生成失败: {error}");
                }

                if (!File.Exists(tempWavFile))
                {
                    return (false, "音频文件生成失败");
                }

                try
                {
                    // 立即读取文件到内存并播放
                    using var fileStream = new FileStream(tempWavFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                    using var memoryStream = new MemoryStream();
                    await fileStream.CopyToAsync(memoryStream);
                    memoryStream.Position = 0;

                    // 立即删除临时文件
                    fileStream.Close();
                    try { File.Delete(tempWavFile); } catch { }

                    // 播放
                    await Task.Run(() =>
                    {
                        using var player = new System.Media.SoundPlayer(memoryStream);
                        player.PlaySync();
                    });

                    return (true, "播放成功");
                }
                catch (Exception ex)
                {
                    try { if (File.Exists(tempWavFile)) File.Delete(tempWavFile); } catch { }
                    throw new Exception($"播放失败: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                return (false, $"播放出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 停止当前播放(暂不支持)
        /// </summary>
        public void Stop()
        {
            // Piper CLI 模式下很难中断播放，需要实现进程管理
            // 这里留作未来改进
        }

        public void Dispose()
        {
            if (!disposed)
            {
                disposed = true;
            }
        }
    }
}
