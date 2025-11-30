using System;
using System.Collections.Generic;
using System.Linq;
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
        private IVoiceService? voiceService;
        private bool isListening = false;
        private string lastClipboardText = "";
        private string lastSpokenText = "";
        private DispatcherTimer? clipboardTimer;
        private bool copyKeyPressed = false;
        private IntPtr keyboardHookId = IntPtr.Zero;
        private LowLevelKeyboardProc? keyboardProc;
        private string currentEngine = "Piper";
        private bool isInitializing = true; // 添加初始化标志
        private bool isSwitchingEngine = false; // 添加切换引擎标志

        public MainWindow()
        {
            InitializeComponent();
            InitializeSpeechSynthesizer();

            AddStatusMessage("程序初始化完成");
        }

        private async void InitializeSpeechSynthesizer()
        {
            try
            {
                isInitializing = true;
                await SwitchToEngine("Piper");
                isInitializing = false;
            }
            catch (Exception ex)
            {
                isInitializing = false;
                MessageBox.Show($"初始化语音合成器失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task SwitchToEngine(string engineName)
        {
            try
            {
                isSwitchingEngine = true;
                
                // 禁用控件，防止用户在切换过程中进行操作
                EngineComboBox.IsEnabled = false;
                VoiceComboBox.IsEnabled = false;
                StartButton.IsEnabled = false;
                TestButton.IsEnabled = false;
                
                AddStatusMessage($"正在切换到 {engineName} 引擎...");

                // 清理旧的服务
                voiceService?.Dispose();
                voiceService = null;

                currentEngine = engineName;

                if (engineName == "Piper")
                {
                    var piperService = new PiperVoiceService();
                    voiceService = piperService;

                    var (isAvailable, checkMessage) = await piperService.IsAvailableAsync();
                    AddStatusMessage(checkMessage);
                    
                    if (!isAvailable)
                    {
                        AddStatusMessage("Piper TTS 未安装，正在尝试安装...");
                        var (success, message) = await piperService.InstallPiperAsync();
                        AddStatusMessage(message);
                        
                        if (!success)
                        {
                            MessageBox.Show($"Piper 安装失败，请手动安装:\npip install piper-tts\n\n{message}", 
                                "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                            // 重新启用控件
                            EngineComboBox.IsEnabled = true;
                            VoiceComboBox.IsEnabled = true;
                            StartButton.IsEnabled = true;
                            TestButton.IsEnabled = true;
                            isSwitchingEngine = false;
                            return;
                        }
                    }

                    // 获取可用的语音
                    AddStatusMessage("正在获取可用语音列表，请稍候...");
                    var voices = await voiceService.GetAvailableVoicesAsync();
                    VoiceComboBox.ItemsSource = voices;

                    if (VoiceComboBox.Items.Count > 0)
                    {
                        VoiceComboBox.SelectedIndex = 0;
                        voiceService.CurrentModel = voices[0];
                    }
                }
                else // Windows TTS
                {
                    voiceService = new WindowsVoiceService();

                    var (isAvailable, checkMessage) = await voiceService.IsAvailableAsync();
                    AddStatusMessage(checkMessage);

                    var voices = await voiceService.GetAvailableVoicesAsync();
                    VoiceComboBox.ItemsSource = voices;

                    if (VoiceComboBox.Items.Count > 0)
                    {
                        VoiceComboBox.SelectedIndex = 0;
                        voiceService.CurrentModel = voices[0];
                    }
                }

                // 设置默认参数
                voiceService.Speed = 1.0;
                voiceService.Volume = 100;
                
                UpdateSettingsDisplay();
                AddStatusMessage($"{voiceService.EngineName} 初始化完成");
                
                // 重新启用控件
                EngineComboBox.IsEnabled = true;
                VoiceComboBox.IsEnabled = true;
                StartButton.IsEnabled = true;
                TestButton.IsEnabled = true;
                
                isSwitchingEngine = false;
            }
            catch (Exception ex)
            {
                AddStatusMessage($"切换引擎失败: {ex.Message}");
                
                // 发生错误时也要重新启用控件
                EngineComboBox.IsEnabled = true;
                VoiceComboBox.IsEnabled = true;
                StartButton.IsEnabled = true;
                TestButton.IsEnabled = true;
                
                isSwitchingEngine = false;
                throw;
            }
        }

        private async void EngineComboBox_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            // 如果正在初始化或正在切换引擎，忽略此事件
            if (isInitializing || isSwitchingEngine)
            {
                return;
            }
            
            if (EngineComboBox.SelectedItem != null)
            {
                var selectedItem = (ComboBoxItem)EngineComboBox.SelectedItem;
                string? engine = selectedItem.Tag?.ToString();
                
                if (!string.IsNullOrEmpty(engine) && engine != currentEngine)
                {
                    if (isListening)
                    {
                        StopButton_Click(null, new RoutedEventArgs());
                    }

                    await SwitchToEngine(engine);
                }
            }
        }

        private void InitializeClipboardTimer()
        {
            if (clipboardTimer == null)
            {
                clipboardTimer = new DispatcherTimer();
                clipboardTimer.Interval = TimeSpan.FromMilliseconds(100);
                clipboardTimer.Tick += ClipboardTimer_Tick;
                AddStatusMessage("剪贴板监听定时器已创建");
            }
        }

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
                if (curModule != null)
                {
                    return SetWindowsHookEx(13, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
                return IntPtr.Zero;
            }
        }

        private IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)0x0100)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                bool ctrlPressed = (GetAsyncKeyState(0x11) & 0x8000) != 0;
                if (ctrlPressed && vkCode == 0x43)
                {
                    if (isListening && !copyKeyPressed)
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
                    keyboardProc = null;
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
            
            MaxTextLengthSlider.Value = maxTextLength;
            MaxTextLengthTextBlock.Text = maxTextLength.ToString();
            
            UpdateSettingsDisplay();
            AddStatusMessage("窗口初始化完成");
        }

        private async void ClipboardTimer_Tick(object? sender, EventArgs e)
        {
            if (!isListening) return;

            try
            {
                string currentText = Clipboard.GetText();

                if (!string.IsNullOrEmpty(currentText) && (currentText != lastClipboardText || copyKeyPressed))
                {
                    var text = ExtractEnglishWords(currentText);
                    if (text.Count() > maxTextLength)
                    {
                        copyKeyPressed = false;
                        lastClipboardText = currentText;
                        AddStatusMessage($"超长文本不与阅读 ( 超过{maxTextLength}个单词，共计{text.Count()}个)");
                        return;
                    }
                    var english = string.Join(" ", text);

                    if (!string.IsNullOrEmpty(english))
                    {
                        lastClipboardText = currentText;

                        bool isRepeatedText = english.Equals(lastSpokenText, StringComparison.OrdinalIgnoreCase);
                        if (isRepeatedText)
                        {
                            AddStatusMessage($"检测到重复内容，跳过播放: {english}");
                        }
                        else
                        {
                            lastSpokenText = english;
                            await SpeakTextAsync(english, originalRate);
                            AddStatusMessage($"正在阅读: {english}");
                        }

                        copyKeyPressed = false;
                    }
                }
            }
            catch (Exception ex)
            {
                AddStatusMessage($"读取剪贴板时出错: {ex.Message}");
            }
        }

        private IEnumerable<string> ExtractEnglishWords(string text)
        {
            if (string.IsNullOrEmpty(text))
                return Enumerable.Empty<string>();

            return Regex.Matches(text, @"[a-zA-Z]+")
                        .Cast<Match>()
                        .Select(match => match.Value);
        }

        private async Task SpeakTextAsync(string text, double speed)
        {
            if (voiceService == null || string.IsNullOrEmpty(text)) return;

            try
            {
                voiceService.Stop();
                voiceService.Speed = speed;

                var (success, message) = await voiceService.SpeakAsync(text);
                
                if (!success)
                {
                    AddStatusMessage($"语音播放失败: {message}");
                }
            }
            catch (Exception ex)
            {
                AddStatusMessage($"语音播放失败: {ex.Message}");
            }
        }

        private async void StartButton_Click(object? sender, RoutedEventArgs e)
        {
            if (voiceService == null)
            {
                MessageBox.Show("语音合成器未初始化，无法开始监听。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (currentEngine == "Piper" && voiceService is PiperVoiceService piperService)
            {
                if (!piperService.IsVoiceDownloaded(voiceService.CurrentModel))
                {
                    AddStatusMessage($"正在下载语音模型: {voiceService.CurrentModel}...");
                    var (success, message) = await piperService.DownloadVoiceAsync(voiceService.CurrentModel);
                    AddStatusMessage(message);
                    
                    if (!success)
                    {
                        MessageBox.Show($"语音模型下载失败，无法开始监听。\n{message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }

            InitializeClipboardTimer();
            RegisterGlobalKeyboardHook();

            isListening = true;
            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            clipboardTimer?.Start();

            lastSpokenText = "";

            AddStatusMessage("开始监听剪贴板...");
        }

        private void StopButton_Click(object? sender, RoutedEventArgs e)
        {
            isListening = false;
            StartButton.IsEnabled = true;
            StopButton.IsEnabled = false;
            clipboardTimer?.Stop();

            voiceService?.Stop();
            UnregisterGlobalKeyboardHook();

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
            string voiceInfo = VoiceDiagnostic.GetVoiceInfo(currentEngine);

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

        private async void VoiceComboBox_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (voiceService != null && VoiceComboBox.SelectedItem != null)
            {
                try
                {
                    string? selectedVoice = VoiceComboBox.SelectedItem.ToString();
                    if (!string.IsNullOrEmpty(selectedVoice))
                    {
                        voiceService.CurrentModel = selectedVoice;
                        
                        if (currentEngine == "Piper" && voiceService is PiperVoiceService piperService)
                        {
                            if (!piperService.IsVoiceDownloaded(selectedVoice))
                            {
                                AddStatusMessage($"正在下载语音模型: {selectedVoice}...");
                                var (success, message) = await piperService.DownloadVoiceAsync(selectedVoice);
                                AddStatusMessage(message);
                            }
                        }
                        
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

        private double originalRate = 1.0;
        private double lostRate = 0.7;
        private int maxTextLength = 4;

        private void SpeedSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (voiceService != null)
            {
                originalRate = SpeedSlider.Value;
                lostRate = originalRate * 0.7;
                voiceService.Speed = originalRate;
                SpeedTextBlock.Text = $"{SpeedSlider.Value:F1}x";
                UpdateSettingsDisplay();
            }
        }

        private void VolumeSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (voiceService != null)
            {
                voiceService.Volume = (int)VolumeSlider.Value;
                VolumeTextBlock.Text = $"{(int)VolumeSlider.Value}%";
                UpdateSettingsDisplay();
            }
        }

        private void MaxTextLengthSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            maxTextLength = (int)MaxTextLengthSlider.Value;
            if (MaxTextLengthTextBlock != null)
            {
                MaxTextLengthTextBlock.Text = maxTextLength.ToString();
            }
            UpdateSettingsDisplay();
        }

        private void UpdateSettingsDisplay()
        {
            if (SettingsTextBlock == null) return;
            
            string engineName = voiceService?.EngineName ?? "未选择";
            string voiceName = VoiceComboBox.SelectedItem?.ToString() ?? "未选择";
            string speed = SpeedTextBlock?.Text ?? "1.0x";
            string volume = VolumeTextBlock?.Text ?? "100%";

            SettingsTextBlock.Text = $"引擎: {engineName} | 语音: {voiceName} | 语速: {speed} | 音量: {volume} | 文本长度限制: {maxTextLength}";
        }

        private void AddStatusMessage(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string fullMessage = $"[{timestamp}] {message}\n";

            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text += fullMessage;

                var scrollViewer = StatusTextBlock.Parent as ScrollViewer;
                if (scrollViewer != null)
                {
                    scrollViewer.ScrollToBottom();
                }
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            clipboardTimer?.Stop();
            voiceService?.Dispose();

            if (isListening)
            {
                UnregisterGlobalKeyboardHook();
            }

            base.OnClosed(e);
        }
    }
}
