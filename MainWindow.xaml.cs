using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Input;
using System.Runtime.InteropServices;

namespace AutoVoice
{
    public partial class MainWindow : Window
    {
        private SpeechSynthesizer? synthesizer;
        private bool isListening = false;
        private string lastClipboardText = "";
        private string lastSpokenText = ""; // 记录上次阅读的文本
        private DispatcherTimer? clipboardTimer;
        private bool copyKeyPressed = false; // 标记是否检测到复制键
        private IntPtr keyboardHookId = IntPtr.Zero; // 全局键盘钩子ID
        private LowLevelKeyboardProc? keyboardProc; // 保持delegate引用

        public MainWindow()
        {
            InitializeComponent();
            InitializeSpeechSynthesizer();
            UpdateSettingsDisplay();

            AddStatusMessage("程序初始化完成");
        }

        private void InitializeSpeechSynthesizer()
        {
            try
            {
                synthesizer = new SpeechSynthesizer();

                // 获取可用的语音
                var voices = synthesizer.GetInstalledVoices();
                VoiceComboBox.ItemsSource = voices.Select(v => v.VoiceInfo.Name).ToList();

                if (VoiceComboBox.Items.Count > 0)
                {
                    VoiceComboBox.SelectedIndex = 0;
                }

                // 设置默认参数
                synthesizer.Rate = 0; // 正常速度
                synthesizer.Volume = 100; // 最大音量
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化语音合成器失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InitializeClipboardTimer()
        {
            if (clipboardTimer == null)
            {
                clipboardTimer = new DispatcherTimer();
                clipboardTimer.Interval = TimeSpan.FromMilliseconds(500); // 每500毫秒检查一次
                clipboardTimer.Tick += ClipboardTimer_Tick;
                AddStatusMessage("剪贴板监听定时器已创建");
            }
        }



        // Windows API 声明
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private void RegisterGlobalKeyboardHook()
        {
            try
            {
                if (keyboardHookId == IntPtr.Zero)
                {
                    // 保持delegate引用，防止被GC回收
                    keyboardProc = KeyboardProc;
                    keyboardHookId = SetHook(keyboardProc);
                    if (keyboardHookId != IntPtr.Zero)
                    {
                        AddStatusMessage("全局键盘钩子注册成功");
                    }
                    else
                    {
                        AddStatusMessage("全局键盘钩子注册失败");
                    }
                }
            }
            catch (Exception ex)
            {
                AddStatusMessage($"注册全局键盘钩子时出错: {ex.Message}");
            }
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(13, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)0x0100) // WM_KEYDOWN
            {
                int vkCode = Marshal.ReadInt32(lParam);

                // 检查是否按下了Ctrl+C
                bool ctrlPressed = (GetAsyncKeyState(0x11) & 0x8000) != 0; // VK_CONTROL
                if (ctrlPressed && vkCode == 0x43) // 'C' key
                {
                    if (isListening && !copyKeyPressed) // 避免重复触发
                    {
                        copyKeyPressed = true;
                        Dispatcher.Invoke(() =>
                        {
                            AddStatusMessage("全局检测到Ctrl+C按键");
                        });
                    }
                }
            }

            return CallNextHookEx(keyboardHookId, nCode, wParam, lParam);
        }

        private void UnregisterGlobalKeyboardHook()
        {
            try
            {
                if (keyboardHookId != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(keyboardHookId);
                    keyboardHookId = IntPtr.Zero;
                    keyboardProc = null; // 释放delegate引用
                    AddStatusMessage("全局键盘钩子已清理");
                }
            }
            catch (Exception ex)
            {
                AddStatusMessage($"清理全局键盘钩子时出错: {ex.Message}");
            }
        }



        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            AddStatusMessage("窗口初始化完成");
        }

        private async void ClipboardTimer_Tick(object? sender, EventArgs e)
        {
            if (!isListening) return;

            try
            {
                string currentText = Clipboard.GetText();

                // 检查是否有新的英文文本或检测到复制键
                if (!string.IsNullOrEmpty(currentText) && (currentText != lastClipboardText || copyKeyPressed))
                {
                    string englishText = ExtractEnglishText(currentText);

                    if (!string.IsNullOrEmpty(englishText))
                    {
                        lastClipboardText = currentText;

                        // 检查是否与上次阅读的文本相同
                        bool isRepeatedText = englishText.Equals(lastSpokenText, StringComparison.OrdinalIgnoreCase);
                        lastSpokenText = englishText;
                        if (isRepeatedText)
                        {
                            // 如果是重复文本，使用低速模式
                            await SpeakTextAsync(englishText, lostRate);
                            AddStatusMessage($"检测到重复内容，使用低速模式重新阅读: {englishText}");
                        }
                        else
                        {
                            // 新文本，使用正常速度
                            await SpeakTextAsync(englishText, originalRate);
                            AddStatusMessage($"正在阅读: {englishText}");
                        }

                        copyKeyPressed = false; // 重置复制键标记
                    }
                }
            }
            catch (Exception ex)
            {
                AddStatusMessage($"读取剪贴板时出错: {ex.Message}");
            }
        }

        private string ExtractEnglishText(string text)
        {
            // 使用正则表达式提取英文单词和句子
            // 匹配英文单词、标点符号和空格
            string pattern = @"[a-zA-Z\s\.,!?;:'""()-]+";
            var matches = Regex.Matches(text, pattern);

            var englishParts = new List<string>();
            foreach (Match match in matches)
            {
                string part = match.Value.Trim();
                if (!string.IsNullOrEmpty(part) && ContainsEnglish(part))
                {
                    englishParts.Add(part);
                }
            }

            return string.Join(" ", englishParts);
        }

        private bool ContainsEnglish(string text)
        {
            // 检查是否包含英文字母
            return text.Any(c => char.IsLetter(c) && c <= 127);
        }

        private async Task SpeakTextAsync(string text, int rate)
        {
            if (synthesizer == null || string.IsNullOrEmpty(text)) return;

            try
            {
                // 停止当前正在播放的语音
                if (synthesizer.State == SynthesizerState.Speaking)
                {
                    synthesizer.SpeakAsyncCancelAll();
                }

                // 异步播放语音
                await Task.Run(() =>
                {
                    // 恢复原始语速设置
                    synthesizer.Rate = rate;
                    synthesizer.SpeakAsync(text);
                });
            }
            catch (Exception ex)
            {
                AddStatusMessage($"语音播放失败: {ex.Message}");
            }
        }
        private void StartButton_Click(object? sender, RoutedEventArgs e)
        {
            if (synthesizer == null)
            {
                MessageBox.Show("语音合成器未初始化，无法开始监听。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 延迟创建定时器
            InitializeClipboardTimer();

            // 注册全局热键
            RegisterGlobalKeyboardHook();

            isListening = true;
            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            clipboardTimer?.Start();

            // 重置文本记录
            lastSpokenText = "";

            AddStatusMessage("开始监听剪贴板...");
        }

        private void StopButton_Click(object? sender, RoutedEventArgs e)
        {
            isListening = false;
            StartButton.IsEnabled = true;
            StopButton.IsEnabled = false;
            clipboardTimer?.Stop();

            // 停止当前播放的语音
            if (synthesizer != null && synthesizer.State == SynthesizerState.Speaking)
            {
                synthesizer.SpeakAsyncCancelAll();
            }

            // 注销全局热键
            UnregisterGlobalKeyboardHook();

            // 重置文本记录
            lastSpokenText = "";

            AddStatusMessage("已停止监听。");
        }

        private async void TestButton_Click(object? sender, RoutedEventArgs e)
        {
            string testText = "Hello, this is a test of the AutoVoice application. The speech synthesis is working correctly.";
            await SpeakTextAsync(testText, originalRate);
            AddStatusMessage("播放测试语音...");
        }

        private void DiagnosticButton_Click(object? sender, RoutedEventArgs e)
        {
            string voiceInfo = VoiceDiagnostic.GetVoiceInfo();

            // 显示诊断信息
            var diagnosticWindow = new Window
            {
                Title = "语音包诊断信息",
                Width = 600,
                Height = 500,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };

            var textBox = new TextBox
            {
                Text = voiceInfo,
                IsReadOnly = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 12
            };

            diagnosticWindow.Content = textBox;
            diagnosticWindow.Show();

            AddStatusMessage("已显示语音包诊断信息");
        }

        private void VoiceComboBox_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (synthesizer != null && VoiceComboBox.SelectedItem != null)
            {
                try
                {
                    string? selectedVoice = VoiceComboBox.SelectedItem.ToString();
                    if (!string.IsNullOrEmpty(selectedVoice))
                    {
                        synthesizer.SelectVoice(selectedVoice);
                        UpdateSettingsDisplay();
                        AddStatusMessage($"已切换到语音: {selectedVoice}");
                    }
                }
                catch (Exception ex)
                {
                    AddStatusMessage($"切换语音失败: {ex.Message}");
                }
            }
        }
        private int originalRate;
        private int lostRate = -5;

        private void SpeedSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (synthesizer != null)
            {
                // 将滑块值转换为语音速率 (-10 到 10)
                originalRate = (int)((SpeedSlider.Value - 1.0) * 10);
                lostRate = (int)(originalRate * 0.5d);
                synthesizer.Rate = originalRate;
                SpeedTextBlock.Text = $"{SpeedSlider.Value:F1}x";
                UpdateSettingsDisplay();
            }
        }

        private void VolumeSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (synthesizer != null)
            {
                synthesizer.Volume = (int)VolumeSlider.Value;
                VolumeTextBlock.Text = $"{(int)VolumeSlider.Value}%";
                UpdateSettingsDisplay();
            }
        }

        private void UpdateSettingsDisplay()
        {
            string voiceName = VoiceComboBox.SelectedItem?.ToString() ?? "未选择";
            string speed = SpeedTextBlock.Text ?? "1.0x";
            string volume = VolumeTextBlock.Text ?? "100%";

            SettingsTextBlock.Text = $"语音: {voiceName} | 语速: {speed} | 音量: {volume}";
        }

        private void AddStatusMessage(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string fullMessage = $"[{timestamp}] {message}\n";

            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text += fullMessage;

                // 自动滚动到底部
                var scrollViewer = StatusTextBlock.Parent as ScrollViewer;
                if (scrollViewer != null)
                {
                    scrollViewer.ScrollToBottom();
                }
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            // 清理资源
            clipboardTimer?.Stop();

            if (synthesizer != null)
            {
                synthesizer.Dispose();
            }

            // 确保注销热键（如果还在监听状态）
            if (isListening)
            {
                UnregisterGlobalKeyboardHook();
            }

            base.OnClosed(e);
        }
    }
}