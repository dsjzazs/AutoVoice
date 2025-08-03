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

namespace AutoVoice
{
    public partial class MainWindow : Window
    {
        private SpeechSynthesizer? synthesizer;
        private bool isListening = false;
        private string lastClipboardText = "";
        private DispatcherTimer? clipboardTimer;

        public MainWindow()
        {
            InitializeComponent();
            InitializeSpeechSynthesizer();
            InitializeClipboardTimer();
            UpdateSettingsDisplay();
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
            clipboardTimer = new DispatcherTimer();
            clipboardTimer.Interval = TimeSpan.FromMilliseconds(500); // 每500毫秒检查一次
            clipboardTimer.Tick += ClipboardTimer_Tick;
        }

        private async void ClipboardTimer_Tick(object? sender, EventArgs e)
        {
            if (!isListening) return;

            try
            {
                string currentText = Clipboard.GetText();
                
                // 检查是否有新的英文文本
                if (!string.IsNullOrEmpty(currentText) && currentText != lastClipboardText)
                {
                    string englishText = ExtractEnglishText(currentText);
                    
                    if (!string.IsNullOrEmpty(englishText))
                    {
                        lastClipboardText = currentText;
                        await SpeakTextAsync(englishText);
                        AddStatusMessage($"正在阅读: {englishText}");
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

        private async Task SpeakTextAsync(string text)
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

            isListening = true;
            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            clipboardTimer.Start();
            
            AddStatusMessage("开始监听剪贴板...");
        }

        private void StopButton_Click(object? sender, RoutedEventArgs e)
        {
            isListening = false;
            StartButton.IsEnabled = true;
            StopButton.IsEnabled = false;
            clipboardTimer.Stop();
            
            // 停止当前播放的语音
            if (synthesizer != null && synthesizer.State == SynthesizerState.Speaking)
            {
                synthesizer.SpeakAsyncCancelAll();
            }
            
            AddStatusMessage("已停止监听。");
        }

        private async void TestButton_Click(object? sender, RoutedEventArgs e)
        {
            string testText = "Hello, this is a test of the AutoVoice application. The speech synthesis is working correctly.";
            await SpeakTextAsync(testText);
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

        private void SpeedSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (synthesizer != null)
            {
                // 将滑块值转换为语音速率 (-10 到 10)
                int rate = (int)((SpeedSlider.Value - 1.0) * 10);
                synthesizer.Rate = rate;
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
            if (clipboardTimer != null)
            {
                clipboardTimer.Stop();
            }
            
            if (synthesizer != null)
            {
                synthesizer.Dispose();
            }
            
            base.OnClosed(e);
        }
    }
} 