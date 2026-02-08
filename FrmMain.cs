using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace DFlasher
{
    public partial class frmMain : Form
    {
        public int MinFreqHz = 1;
        public int MaxFreqHz = 200;

        private const string CurrentDigitCmd = "D";
        private const string BrightnessCmd = "B";
        private const string CurrentFreqCmd = "F";
        private const string IncCurrentFreqCmd = "F+";
        private const string DecCurrentFreqCmd = "F-";

        private static string configFile = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), @"Config\settings.config");
        Configuration Cfg = new Configuration();


        CommunicationManager comm = new CommunicationManager();

        // Поля хука
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc _proc;       // держим делегат, чтобы GC не собрал
        private IntPtr _hookID = IntPtr.Zero;

        // WinAPI
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        // ===== Хук =====
        private static IntPtr _hook = IntPtr.Zero;
        private LowLevelKeyboardProc _hookProc; // удерживаем делегат, чтобы GC не собрал

        //для mmTimer
        public int mTimerId;
        private TimerEventHandler mHandler;  // NOTE: declare at class scope so garbage collector doesn't release it!!!
        private delegate void TimerEventHandler(int id, int msg, IntPtr user, int dw1, int dw2);
        private const int TIME_PERIODIC = 1;
        private const int EVENT_TYPE = TIME_PERIODIC;// + 0x100;  // TIME_KILL_SYNCHRONOUS causes a hang ?!
        [DllImport("winmm.dll")]
        private static extern int timeSetEvent(int delay, int resolution, TimerEventHandler handler, IntPtr user, int eventType);
        [DllImport("winmm.dll")]
        private static extern int timeKillEvent(int id);
        [DllImport("winmm.dll")]
        private static extern int timeBeginPeriod(int msec);
        [DllImport("winmm.dll")]
        private static extern int timeEndPeriod(int msec);
        //
        [DllImport("Kernel32.dll")]
        private static extern bool QueryPerformanceCounter(out long lpPerformanceCount);
        [DllImport("Kernel32.dll")]
        private static extern bool QueryPerformanceFrequency(out long lpFrequency);
        public enum TPresentationState { None, Presenting, WaitingInput, Stopped };
        public enum WS { wsStoped, wsFreqSearch, wsWaiting, wsFreqFound, wsFreqSearchByStep, wsReset };
        public WS WrkState = WS.wsStoped;

        private TPresentationState _seqState = TPresentationState.None;
        private bool _seqInputStarted = false; // пользователь уже начал ввод 
        private long _seqFirstKeyPc = 0;   // момент первой реакции

        private int _seqIndex = 0;
        private long _seqStimulusStartPc = 0;   //для отслеживания Duration текущего стимула

        public long Pf;
        public long Pc;
        public long LastPc;
        public long LastPc2;
        public long StartPc;

        private volatile bool _sessionRunning = false;

        private SynchronizationContext _ui;

        private long _intervalTicks; // период обновления для отображения времени
        long StepTimeInTicks;

        private int CurrentFreq;
        private int CurrentIteration;
        private int CurrentThreshold;
        private int IteratinsCount;
        private int NeedRandomDigits;
        private static int NeedSwapDigits;
        private double AvgThreshold = 0;
        private List<double> _allThresholds = new List<double>();
        // направление «ползания»: -1 — вниз, +1 — вверх
        private int _sweepDir = -1;          // начнём с понижения
        //private bool _jumpUpNext = true;     // первый скачок — вверх
        private Random _rnd;
        private HashSet<int> _excludedDigits = null;

        // ===== Данные для анализа =====
        private readonly object _lock = new object();
        private readonly List<int> _freqAtKeypress = new List<int>();
        private readonly List<double> _pcAtKeypress = new List<double>();
        private readonly List<bool> _isCorrect = new List<bool>();

        // буфер ввода цифр в состоянии wsWaiting
        private readonly StringBuilder _digitBuf = new StringBuilder(2);

        private int Brightness = 5;
        private int TwoDigitNumber = 55;

        private DXSoundPlayer _SndPlayer;

        public frmMain()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(configFile));

            QueryPerformanceFrequency(out Pf);
            _intervalTicks = Pf / 100; // 10 мс
            QueryPerformanceCounter(out LastPc);

            timeBeginPeriod(1);
            mHandler = new TimerEventHandler(TmrCallbckForMainProcessing);
            mTimerId = timeSetEvent(1, 1, mHandler, IntPtr.Zero, EVENT_TYPE);

            _hookProc = KeyboardHookProc;
            var hMod = GetModuleHandle(null);
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, hMod, 0);

            InitializeComponent();

            lstBxStimWorkSequence.DrawMode = DrawMode.OwnerDrawFixed;
            lstBxStimWorkSequence.DrawItem += lstBxStimWorkSequence_DrawItem;

            _ui = SynchronizationContext.Current; // 

            _sessionRunning = false;
            StartPc = 0;
            txtBxExpTime.Text = "00:00:00.000";   // напрямую, без Post

            InitRandom();

            this.KeyPreview = true;
            tabCtrlStimulusSettings.SelectedTab = tabSequential;

        }
        private void PrepareExcludedDigits(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                _excludedDigits = null;
                return;
            }

            _excludedDigits = new HashSet<int>();

            char[] sep = new[] { ',', ';', ' ', '\t', '\r', '\n' };
            var tokens = raw.Split(sep, StringSplitOptions.RemoveEmptyEntries);

            foreach (var t in tokens)
            {
                if (int.TryParse(t.Trim(), out int n))
                {
                    if (n >= 10 && n <= 99)
                        _excludedDigits.Add(n);
                }
            }

            // если исключили вообще всё — отключаем исключения
            if (_excludedDigits.Count >= 90)
                _excludedDigits = null;
        }
        private int GenerateAllowedTwoDigitNumber()
        {
            if (_excludedDigits == null)
                return _rnd.Next(10, 100);

            int n;

            do
            {
                n = _rnd.Next(10, 100);
            }
            while (_excludedDigits.Contains(n));

            return n;
        }
        private void SendRandomDigits()
        {
            TwoDigitNumber = GenerateAllowedTwoDigitNumber();//_rnd.Next(10, 100);
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
        }
        private void ProcessSequentialTimer()
        {
            if (_seqState == TPresentationState.None ||
                _seqState == TPresentationState.Stopped)
                return;

            if (Cfg.Stimulus == null ||
                Cfg.Stimulus.Count == 0 ||
                _seqIndex < 0 ||
                _seqIndex >= Cfg.Stimulus.Count)
            {
                QueryPerformanceCounter(out long pcNow);
                StopSequential(pcNow, "invalid_index_or_empty_list", false);
                return;
            }

            // если человек уже начал ввод — таймер больше не действует
            if (_seqInputStarted)
                return;

            QueryPerformanceCounter(out Pc);

            var s = Cfg.Stimulus[_seqIndex];

            long elapsedMs = (Pc - _seqStimulusStartPc) * 1000 / Pf;

            if (_seqState != TPresentationState.Presenting)
                return;

            if (elapsedMs < s.Duration)
                return;

            // ===== TIMEOUT =====

            if (s.NeedReaction)
            {
                comm.WriteData(CurrentFreqCmd + "196");
                CurrentFreq = 196;

                _seqState = TPresentationState.WaitingInput;
                _digitBuf.Clear();

                Logger.WriteLog(Pc,
                    "stimulus_timeout",
                    $"idx={_seqIndex}");

                return;
            }

            // без реакции → следующий

            _seqIndex++;

            if (_seqIndex >= Cfg.Stimulus.Count)
            {
                StopSequential(Pc, "completed_all_stimuli", false);
                return;
            }

            var next = Cfg.Stimulus[_seqIndex];

            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(next.Digits));
            comm.WriteData(CurrentFreqCmd + next.Frequency);
            CurrentFreq = next.Frequency;

            QueryPerformanceCounter(out _seqStimulusStartPc);

            _ui.Post(_ =>
            {
                txtBxCurItertn.Text = (_seqIndex + 1).ToString();
                lstBxStimWorkSequence.SelectedIndex = -1;
                lstBxStimWorkSequence.SelectedIndex = _seqIndex;
                lstBxStimWorkSequence.Invalidate();
            }, null);

            _seqState = TPresentationState.Presenting;
            _seqInputStarted = false;
            _seqFirstKeyPc = 0;
            _digitBuf.Clear();

            Logger.WriteLog(_seqStimulusStartPc,
                "stimulus_start",
                $"idx={_seqIndex};digits={next.Digits};freq={next.Frequency};dur={next.Duration};needReaction={(next.NeedReaction ? 1 : 0)}");
        }
        private void UpdateExpTime(long pcNow)
        {
            if (!_sessionRunning) return;
            if (StartPc <= 0) return;
            if (pcNow < StartPc) return;

            if (pcNow < LastPc + _intervalTicks) return;

            double elapsedSeconds = (double)(pcNow - StartPc) / Pf;
            string formatted = TimeSpan.FromSeconds(elapsedSeconds).ToString(@"hh\:mm\:ss\.fff");

            _ui.Post(_ => txtBxExpTime.Text = formatted, null);
            LastPc = pcNow;
        }
        private void TmrCallbckForMainProcessing(int id, int msg, IntPtr user, int dw1, int dw2)
        {
            QueryPerformanceCounter(out Pc);

            if (Cfg.Mode == ExperimentMode.Sequential)
            {
                ProcessSequentialTimer();
                UpdateExpTime(Pc);
                return;
            }

            // ===== Staircase как было =====
            if (WrkState != WS.wsStoped)
            {
                if (WrkState == WS.wsFreqSearch)
                {
                    if (Pc >= LastPc2 + StepTimeInTicks)
                    {
                        if (_sweepDir < 0)
                        {
                            if (CurrentFreq > MinFreqHz)
                            {
                                if (!comm.WriteData(DecCurrentFreqCmd))
                                {
                                    SafeStopExperiment("Потеря связи с устройством");
                                    return;
                                }
                                CurrentFreq--;
                                if (NeedRandomDigits == 1) SendRandomDigits();
                            }
                            else
                            {
                                if (_SndPlayer != null) _SndPlayer.StopAll();
                            }
                        }
                        else
                        {
                            if (CurrentFreq < Cfg.StartingFreq)
                            {
                                if (!comm.WriteData(IncCurrentFreqCmd))
                                {
                                    SafeStopExperiment("Потеря связи с устройством");
                                    return;
                                }
                                CurrentFreq++;
                                if (NeedRandomDigits == 1) SendRandomDigits();
                            }
                            else
                            {
                                if (_SndPlayer != null) _SndPlayer.StopAll();
                            }
                        }

                        LastPc2 = Pc;
                    }
                }

                UpdateExpTime(Pc);
            }
        }
        private void SafeStopExperiment(string reason)
        {        
            // Сохраняем состояние
            bool wasRunning = (WrkState != WS.wsStoped);

            // Останавливаем эксперимент
            WrkState = WS.wsStoped;

            _sessionRunning = false;

            // Останавливаем звуки
            if (_SndPlayer != null)
            {
                try { _SndPlayer.StopAll(); } catch { }
            }

            // Обновляем UI безопасно
            _ui.Post(_ =>
            {
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                lstBxStimWorkSequence.Enabled = true;
                txtBxCurFreq.Text = "---";

                // Показываем сообщение только если эксперимент был активен
                if (wasRunning)
                {
                    MessageBox.Show(this,
                        $"{reason}\nЭксперимент остановлен.\n\n" +
                        "Проверьте подключение устройства и нажмите 'Открыть'.",
                        "Ошибка связи",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

            }, null);

            // Логируем остановку
            QueryPerformanceCounter(out long pcNow);
            Logger.WriteLog(pcNow, "experiment_stopped", $"reason={reason}");
        }

        // helper для извлечения цифры из VK (верхний ряд и NumPad)
        private static bool TryGetDigitFromVk(int vk, out int digit)
        {
            // верхний ряд 0..9
            if (vk >= (int)Keys.D0 && vk <= (int)Keys.D9)
            {
                digit = vk - (int)Keys.D0;    // 0..9
                return true;
            }
            // NumPad 0..9
            if (vk >= (int)Keys.NumPad0 && vk <= (int)Keys.NumPad9)
            {
                digit = vk - (int)Keys.NumPad0; // 0..9
                return true;
            }

            digit = 0;
            return false;
        }
        private void HandleSequentialKeypress(int vk, long pcNow)
        {
            if (_seqState != TPresentationState.Presenting &&
                _seqState != TPresentationState.WaitingInput)
                return;

            // первая реакция — фиксируем RT
            if (!_seqInputStarted)
            {
                _seqInputStarted = true;
                _seqFirstKeyPc = pcNow;
            }

            // цифры
            if (TryGetDigitFromVk(vk, out int digit))
            {
                if (_digitBuf.Length < 2)
                    _digitBuf.Append((char)('0' + digit));
                return;
            }

            // Enter
            if (vk != 0x0D)
                return;

            if (_digitBuf.Length != 2)
                return;

            int entered = int.Parse(_digitBuf.ToString());
            _digitBuf.Clear();

            int target = Cfg.Stimulus[_seqIndex].Digits;
            bool correct = (entered == target);

            // RT
            string rtStr = "n/a";

            if (_seqState == TPresentationState.Presenting &&
                _seqFirstKeyPc > 0)
            {
                long rtMs = (_seqFirstKeyPc - _seqStimulusStartPc) * 1000 / Pf;
                rtStr = rtMs.ToString();
            }

            // ===== LOG =====
            Logger.WriteLog(pcNow,
                "stimulus_answer",
                $"idx={_seqIndex};input={entered};target={target};correct={(correct ? 1 : 0)};rtMs={rtStr}");

            // ===== следующий =====

            _seqIndex++;

            if (_seqIndex >= Cfg.Stimulus.Count)
            {
                StopSequential(pcNow, "completed_all_stimuli", false);
                return;
            }

            var s = Cfg.Stimulus[_seqIndex];

            QueryPerformanceCounter(out _seqStimulusStartPc);

            _seqState = TPresentationState.Presenting;
            _seqInputStarted = false;
            _seqFirstKeyPc = 0;
            _digitBuf.Clear();

            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(s.Digits));
            comm.WriteData(CurrentFreqCmd + s.Frequency);
            CurrentFreq = s.Frequency;

            _ui.Post(_ =>
            {
                txtBxCurItertn.Text = (_seqIndex + 1).ToString();
                lstBxStimWorkSequence.SelectedIndex = -1;
                lstBxStimWorkSequence.SelectedIndex = _seqIndex;
                lstBxStimWorkSequence.Invalidate();
            }, null);

            Logger.WriteLog(_seqStimulusStartPc,
                "stimulus_start",
                $"idx={_seqIndex};digits={s.Digits};freq={s.Frequency};dur={s.Duration};needReaction={(s.NeedReaction ? 1 : 0)}");
        }
        // ==== Глобальный хук: ловим любую клавишу, цифры и Enter ====
        private IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            const int WM_KEYDOWN_LOCAL = 0x0100;
            if (wParam != (IntPtr)WM_KEYDOWN_LOCAL)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            if (!comm.ComPortIsOpen)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            // ВАЖНО: Sequential разрешаем даже если WrkState == wsStoped
            bool seqRunning = (Cfg.Mode == ExperimentMode.Sequential) &&
                              (_seqState != TPresentationState.None) &&
                              (_seqState != TPresentationState.Stopped);

            if (!seqRunning && WrkState == WS.wsStoped)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            KBDLLHOOKSTRUCT k = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            int vk = k.vkCode;

            QueryPerformanceCounter(out long pcNow);

            if (Cfg.Mode == ExperimentMode.Sequential)
            {
                HandleSequentialKeypress(vk, pcNow);
                return CallNextHookEx(_hook, nCode, wParam, lParam);
            }

            // ===== Staircase логика как было =====
            double msFromStart = (pcNow - StartPc) * 1000.0 / Pf;
            int freqNow = CurrentFreq;

            if (WrkState != WS.wsWaiting)
            {
                WrkState = WS.wsWaiting;
                lock (_lock)
                {
                    _freqAtKeypress.Add(freqNow);
                    _pcAtKeypress.Add(msFromStart);
                    _isCorrect.Add(false);
                }
                Logger.WriteLog(pcNow, "keypress", "vk=" + vk.ToString());

                if (_SndPlayer != null)
                {
                    _SndPlayer.Stop("shep_up");
                    _SndPlayer.Stop("shep_down");
                }
            }

            if (WrkState == WS.wsWaiting)
            {
                int digit;
                if (TryGetDigitFromVk(vk, out digit))
                {
                    if (_digitBuf.Length < 2)
                        _digitBuf.Append((char)('0' + digit));
                }
                else if (vk == 0x0D) // Enter
                {
                    if (_digitBuf.Length == 2)
                    {
                        int entered = int.Parse(_digitBuf.ToString());
                        _digitBuf.Clear();

                        bool ok = (entered == TwoDigitNumber);

                        if (ok)
                        {
                            lock (_lock)
                            {
                                if (_isCorrect.Count > 0)
                                    _isCorrect[_isCorrect.Count - 1] = true;
                            }

                            CurrentIteration++;
                            CurrentThreshold = freqNow;

                            _allThresholds.Add(CurrentThreshold);
                            AvgThreshold = _allThresholds.Average();

                            double sigma = ComputeSigmaHz();

                            _ui.Post(_ =>
                            {
                                txtBxCurItertn.Text = CurrentIteration.ToString();
                                txtBxCurThresshld.Text = CurrentThreshold.ToString("0.###");
                                txtBxAvrgThresshld.Text = AvgThreshold.ToString("0.###");
                                txtBxCurSigma.Text = sigma.ToString("0.###");
                            }, null);

                            if (CurrentIteration >= IteratinsCount && sigma <= Cfg.StdDevThreshold)
                            {
                                string details = string.Format(
                                    "input={0};target={1};correct=1;CurFreq={2};σ={3:0.###}",
                                    entered, TwoDigitNumber, freqNow, sigma);

                                Logger.WriteLog(pcNow, "answer", details);

                                WrkState = WS.wsStoped;
                                _sessionRunning = false;
                                Logger.WriteLog(pcNow, "stop",
                                    string.Format("σ-threshold reached;freq={0};σ={1:0.###}", freqNow, sigma));

                                comm.WriteData(CurrentFreqCmd + Cfg.StartingFreq.ToString());
                                Logger.Flush();

                                if (_SndPlayer != null)
                                {
                                    _SndPlayer.Stop("shep_up");
                                    _SndPlayer.Stop("shep_down");
                                    _SndPlayer.Play("success_pabam", false);
                                }

                                _ui.Post(_ =>
                                {
                                    btnStart.Enabled = true;
                                    btnStop.Enabled = false;
                                    lstBxStimWorkSequence.Enabled = true;
                                }, null);
                            }
                            else
                            {
                                string details = string.Format(
                                    "input={0};target={1};correct=1;CurFreq={2};σ={3:0.###}",
                                    entered, TwoDigitNumber, freqNow, sigma);

                                Logger.WriteLog(pcNow, "answer", details);

                                int oldDir = _sweepDir;
                                int newStartFreq = CurrentFreq + oldDir * Cfg.StepJumpFreq;

                                if (newStartFreq < 1) newStartFreq = 1;
                                if (newStartFreq > Cfg.StartingFreq) newStartFreq = Cfg.StartingFreq;

                                CurrentFreq = newStartFreq;
                                comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());

                                SendRandomDigits();

                                _sweepDir = -oldDir;
                                string workDir = (_sweepDir < 0 ? "down" : "up");

                                if (_SndPlayer != null)
                                {
                                    if (_sweepDir < 0)
                                    {
                                        _SndPlayer.Stop("shep_up");
                                        _SndPlayer.Play("shep_down", true);
                                    }
                                    else
                                    {
                                        _SndPlayer.Stop("shep_down");
                                        _SndPlayer.Play("shep_up", true);
                                    }
                                }

                                Logger.WriteLog(pcNow, "jump",
                                    string.Format("NewStartFreq={0};dir={1}", CurrentFreq, workDir));

                                LastPc2 = pcNow;
                                WrkState = WS.wsFreqSearch;
                            }
                        }
                        else
                        {
                            string details = string.Format(
                                "input={0};target={1};correct=0;CurFreq={2}",
                                entered, TwoDigitNumber, freqNow);

                            Logger.WriteLog(pcNow, "answer", details);

                            int oldDir = _sweepDir;
                            int newStartFreq = CurrentFreq + oldDir * Cfg.StepJumpFreq;

                            if (newStartFreq < 1) newStartFreq = 1;
                            if (newStartFreq > Cfg.StartingFreq) newStartFreq = Cfg.StartingFreq;

                            CurrentFreq = newStartFreq;
                            comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());

                            SendRandomDigits();

                            _sweepDir = -oldDir;
                            string workDir = (_sweepDir < 0 ? "down" : "up");

                            if (_SndPlayer != null)
                            {
                                if (_sweepDir < 0)
                                {
                                    _SndPlayer.Stop("shep_up");
                                    _SndPlayer.Play("shep_down", true);
                                }
                                else
                                {
                                    _SndPlayer.Stop("shep_down");
                                    _SndPlayer.Play("shep_up", true);
                                }
                            }

                            Logger.WriteLog(pcNow, "jump",
                                string.Format("NewStartFreq={0};dir={1}", CurrentFreq, workDir));

                            LastPc2 = pcNow;
                            WrkState = WS.wsFreqSearch;
                        }
                    }
                }
            }

            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        // ==== Подсчёт σ по корректным ====
        private double ComputeSigmaHz()
        {
            List<double> vals = new List<double>();
            lock (_lock)
            {
                for (int i = 0; i < _isCorrect.Count; i++)
                    if (_isCorrect[i]) vals.Add(_freqAtKeypress[i]);
            }
            if (vals.Count < 2) return double.PositiveInfinity;

            double mean = 0;
            foreach (var v in vals) mean += v;
            mean /= vals.Count;

            double s2 = 0;
            foreach (var v in vals) { double d = v - mean; s2 += d * d; }
            s2 /= (vals.Count - 1);
            return Math.Sqrt(s2);
        }
        // ==== инициализация рандома ====
        private void InitRandom()
        {
            byte[] seed = new byte[4];
            using (var rng = RandomNumberGenerator.Create())
                rng.GetBytes(seed);
            int s = BitConverter.ToInt32(seed, 0);
            _rnd = new Random(s);

            // фиксируем сид в лог
            //Logger.WriteLog(StartPc, "seed_init", $"seed={s}");
        }
        // ==== хелпер: логическое → для железа (перевернуть цифры) ====
        private static int ToHardwareDigit(int n)
        {
            if (NeedSwapDigits == 1)
            {
                // на всякий случай нормализуем в диапазон 0..99
                if (n < 0) n = 0;
                if (n > 99) n = n % 100;

                int tens = n / 10;
                int ones = n % 10;
                return ones * 10 + tens; // 13 -> 31, 07 -> 70, 5 -> 50
            }
            else
            {
                return n;
            }
        }
        private void AutoConnectCh340()
        {
            // ищем все CH340
            var chPorts = SerialHelpers.FindCh340Ports();

            if (chPorts.Length == 0)
            {
                // не нашли — остаёмся в ручном режиме
                if (cboPort.Items.Count > 0 && cboPort.SelectedIndex < 0)
                    cboPort.SelectedIndex = 0;

                MessageBox.Show(this,
                    "Устройство \"USB-SERIAL CH340\" не найдено.\n" +
                    "Выберите COM-порт вручную и нажмите \"Открыть\".",
                    "CH340 не найден",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // берём первый найденный порт CH340
            string port = chPorts[0];

            // убеждаемся, что он есть в списке
            if (!cboPort.Items.Contains(port))
                cboPort.Items.Add(port);

            cboPort.SelectedItem = port;

            // пробуем открыть его с текущими настройками
            TryOpenSelectedPort();
        }
        /// <summary>
        /// Автоподключение к Bluetooth-COM порту (HC-06).
        /// Ищет Bluetooth-порт через SerialHelpers.FindBluetoothOutPort()
        /// и пытается сразу открыть его с текущими настройками.
        /// </summary>
        private void AutoConnectBluetooth()
        {
            string btPort = SerialHelpers.FindBluetoothOutPort();

            if (string.IsNullOrEmpty(btPort))
            {
                // Не нашли BT-порт — остаёмся в ручном режиме
                if (cboPort.Items.Count > 0 && cboPort.SelectedIndex < 0)
                    cboPort.SelectedIndex = 0;

                MessageBox.Show(this,
                    "Bluetooth COM-порт (HC-06) не найден.\n" +
                    "Убедитесь, что модуль спарен с Windows,\n" +
                    "и выберите COM-порт вручную.",
                    "Bluetooth-порт не найден",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Добавим в список, если его там ещё нет
            if (!cboPort.Items.Contains(btPort))
                cboPort.Items.Add(btPort);

            cboPort.SelectedItem = btPort;

            // пробуем открыть его
            TryOpenSelectedPort();
        }
        private void TryOpenSelectedPort()
        {
            // настройки порта — как в твоём cmdOpen_Click
            comm.Parity = cboParity.Text;
            comm.StopBits = cboStop.Text;
            comm.DataBits = cboData.Text;
            comm.BaudRate = cboBaud.Text;
            comm.DisplayWindow = rtbDisplay;
            comm.DisplayCtrl1 = null;
            comm.DisplayCtrl2 = null;
            comm.DisplayCtrl3 = null;
            comm.PortName = cboPort.Text;

            if (comm.OpenPort())
            {
                cmdOpen.Enabled = false;
                cmdClose.Enabled = true;
                cmdSend.Enabled = true;
            }
            else
            {
                // если не удалось открыть — просто остаёмся в ручном режиме
                cmdOpen.Enabled = true;
                cmdClose.Enabled = false;
                cmdSend.Enabled = false;

                MessageBox.Show(this,
                    $"Не удалось открыть порт {cboPort.Text}.\n" +
                    "Выберите порт вручную и нажмите \"Открыть\".",
                    "Ошибка открытия порта",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            Cfg = Configuration.Load(configFile);

            CurrentFreq = Cfg.StartingFreq;
            NeedSwapDigits = Cfg.NeedSwapDigits;
            NeedRandomDigits = Cfg.NeedRandomDigits;

            LoadComPortValues();
            SetComPortDefaults();
            SetControlsStates();

           // AutoConnectCh340();

            if (!comm.ComPortIsOpen)
            {
                AutoConnectBluetooth();
            }

            // Инициализация шепард-плеера
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory + "Data";
                string upPath = System.IO.Path.Combine(baseDir, "shepard_up_loop.wav");
                string downPath = System.IO.Path.Combine(baseDir, "shepard_down_loop.wav");
                string succPath = System.IO.Path.Combine(baseDir, "success_pabam.wav");

                _SndPlayer = new DXSoundPlayer(this.Handle);
                _SndPlayer.LoadWav("shep_up", upPath);
                _SndPlayer.LoadWav("shep_down", downPath);
                _SndPlayer.LoadWav("success_pabam", succPath);


            }
            catch (Exception ex)
            {
                // Если нет файлов или DirectSound не завёлся — просто пишем в лог/RTB и живём без звука
                comm.DisplayWindow.AppendText("DXSoundPlayer init error: " + ex.Message + Environment.NewLine);
            }
        }
        private void LoadComPortValues()
        {
            comm.SetPortNameValues(cboPort);
            comm.SetParityValues(cboParity);
            comm.SetStopBitValues(cboStop);
        }
        private void SetComPortDefaults()
        {
            cboPort.SelectedIndex = cboPort.Items.Count - 1;
            int idx = cboBaud.FindStringExact("115200");
            if (idx >= 0)
                cboBaud.SelectedIndex = idx;
            //cboBaud.SelectedText = "115200";
            cboParity.SelectedIndex = 0;
            cboStop.SelectedIndex = 1;
            cboData.SelectedIndex = 1;
        }
        private void RefreshStimulusList()
        {
            // сохраняем текущие выделенные индексы
            int selCount = lstBxStimWorkSequence.SelectedIndices.Count;
            int[] selected = new int[selCount];
            for (int i = 0; i < selCount; i++)
                selected[i] = lstBxStimWorkSequence.SelectedIndices[i];

            GVars.LockCtrl = true;

            lstBxStimWorkSequence.BeginUpdate();
            try
            {
                lstBxStimWorkSequence.Items.Clear();
                if (Cfg.Stimulus == null) return;

                for (int i = 0; i < Cfg.Stimulus.Count; i++)
                {
                    var s = Cfg.Stimulus[i];
                    int idx = i + 1;

                    lstBxStimWorkSequence.Items.Add(
                        $"{idx}. Digit={s.Digits}; Freq={s.Frequency}; Durtn={s.Duration}; NeedReact={(s.NeedReaction ? "true" : "false")}"
                    );
                }

                // восстановим выделение
                for (int i = 0; i < selected.Length; i++)
                {
                    int idx = selected[i];
                    if (idx >= 0 && idx < lstBxStimWorkSequence.Items.Count)
                        lstBxStimWorkSequence.SetSelected(idx, true);
                }
            }
            finally
            {
                lstBxStimWorkSequence.EndUpdate();
                GVars.LockCtrl = false;
            }
        }
        private void SetControlsStates()
        {
            GVars.LockCtrl = true;

            RefreshStimulusList();

            rdoHex.Checked = true;
            // rdoText.Checked = true;
            cmdSend.Enabled = false;
            cmdClose.Enabled = false;

            numUpDwnStartingFreq.Value = Cfg.StartingFreq;
            numUpDwnStepFreqHz.Value = Cfg.StepJumpFreq;
            numUpDwnTimeStepFreqChng.Value = Cfg.TimeStepFreqChngMs;
            numUpDwnStepClarifyThrshld.Value = Cfg.StepClarifyThrshldHz;
            numUpDwnStdDevIn3Itrtn.Value = Cfg.StdDevThreshold;
            numUpDwnBrightness.Value = Cfg.Brightness;

            numUpDwnItrtnCount.Value = Cfg.IteratinsCount;

            // Режим эксперимента
            rdoBtnStaircase.Checked = (Cfg.Mode == ExperimentMode.Staircase);
            rdoBtnSequential.Checked = (Cfg.Mode == ExperimentMode.Sequential);

            chckBxNeedRndDigits.Checked = Convert.ToBoolean(Cfg.NeedRandomDigits);

            // Sequential настройки (только нужные)
            
            numUpDwnStartingFreqSequentialMode.Value = Cfg.SequentialStartingFreq;
            numUpDwnTrialCount.Value = Cfg.SequentialTrials;
            numUpDwnBrightnessSequentialMode.Value = Cfg.SequentialBrightness;
            numUpDwnTrialCount.Value = lstBxStimWorkSequence.Items.Count; 

            txtBxDigitsForExclusion.Text = Cfg.DigitsForExclusion;

            GVars.LockCtrl = false;
        }

        private void GetControlsStates()
        {
            if (rdoBtnStaircase.Checked)
                Cfg.Mode = ExperimentMode.Staircase;
            else if (rdoBtnSequential.Checked)
                Cfg.Mode = ExperimentMode.Sequential;
            // Staircase настройки
            Cfg.StartingFreq = (int)numUpDwnStartingFreq.Value;
            Cfg.StepJumpFreq = (int)numUpDwnStepFreqHz.Value;
            Cfg.TimeStepFreqChngMs = (int)numUpDwnTimeStepFreqChng.Value;
            Cfg.StepClarifyThrshldHz = (int)numUpDwnStepClarifyThrshld.Value;
            Cfg.StdDevThreshold = (int)numUpDwnStdDevIn3Itrtn.Value;
            Cfg.Brightness = (int)numUpDwnBrightness.Value;
            Cfg.IteratinsCount = (int)numUpDwnItrtnCount.Value;
            // Sequential настройки
            Cfg.SequentialStartingFreq = (int)numUpDwnStartingFreqSequentialMode.Value;
            Cfg.SequentialTrials = (int)numUpDwnTrialCount.Value;
            Cfg.SequentialBrightness = (int)numUpDwnBrightnessSequentialMode.Value;

            Cfg.NeedRandomDigits = Convert.ToInt32(chckBxNeedRndDigits.Checked);

            Cfg.DigitsForExclusion = txtBxDigitsForExclusion.Text.Trim();
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_SndPlayer != null)
            {
                _SndPlayer.Dispose();
                _SndPlayer = null;
            }

            if (comm.ComPortIsOpen)
                comm.WriteData(CurrentFreqCmd + 50);

            int i = 0;
            while (i < 20)
            {
                Application.DoEvents();
                Thread.Sleep(80);
                i++;
            }

            Logger.Flush();

            if (_hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hook);
                _hook = IntPtr.Zero;
            }
            _proc = null;

            GetControlsStates();

            Cfg.Save(configFile);

            comm.ClosePort();

            WrkState = WS.wsStoped;

            i = 0;
            while (i < 10)
            {
                Application.DoEvents();
                Thread.Sleep(80);
                i++;
            }

            timeEndPeriod(1);
            int err = timeKillEvent(mTimerId);
            mTimerId = 0;
        }

        private void cmdOpen_Click(object sender, EventArgs e)
        {
            comm.Parity = cboParity.Text;
            comm.StopBits = cboStop.Text;
            comm.DataBits = cboData.Text;
            comm.BaudRate = cboBaud.Text;
            comm.DisplayWindow = rtbDisplay;
            comm.DisplayCtrl1 = null;
            comm.DisplayCtrl2 = null;
            comm.DisplayCtrl3 = null;
            comm.PortName = cboPort.Text;
            comm.CurrentTransmissionType = CommunicationManager.TransmissionType.Text;
            if (comm.OpenPort())
            {
                // Авто-запрос статуса у ардуины
                // WriteData сам добавит '\n', если его нет
                comm.WriteData("S");

                cmdOpen.Enabled = false;
                cmdClose.Enabled = true;
                cmdSend.Enabled = true;
            }
            else
            {
                // Порт не открылся — не даём пользователю нажимать send/close
                cmdOpen.Enabled = true;
                cmdClose.Enabled = false;
                cmdSend.Enabled = false;
            }
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            if (comm.ClosePort())
            {
                cmdOpen.Enabled = true;
                cmdClose.Enabled = false;
                cmdSend.Enabled = false;
            }
        }

        private void btHarshCtrlFaster_Click(object sender, EventArgs e)
        {
            if (!comm.WriteData(IncCurrentFreqCmd))
            {
                MessageBox.Show("Не удалось отправить команду. Проверьте связь.",
                               "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btHarshCtrlSlower_Click(object sender, EventArgs e)
        {
            if (!comm.WriteData(DecCurrentFreqCmd))
            {
                MessageBox.Show("Не удалось отправить команду. Проверьте связь.",
                               "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void txtBxBrightness_TextChanged(object sender, EventArgs e)
        {
            int bValue = int.TryParse(txtBxBrightness.Text, out var tmp) ? tmp : 0;
            //if (comm.ComPortIsOpen) comm.WriteData(BrightnessCmd + bValue.ToString());
            if (!comm.WriteData(BrightnessCmd + bValue.ToString()))
            {
                // Можно показать сообщение или просто сбросить фокус
                this.ActiveControl = null;
            }
        }

        private void txtBxCurDigit_TextChanged(object sender, EventArgs e)
        {
            int dValue = int.TryParse(txtBxCurDigit.Text, out var tmp) ? tmp : 0;

            //if (comm.ComPortIsOpen) comm.WriteData(CurrentDigitCmd + ToHardwareDigit(dValue));
            if (!comm.WriteData(CurrentDigitCmd + ToHardwareDigit(dValue)))
            {
                this.ActiveControl = null;
            }
        }

        private void txtBxCurFreq_TextChanged(object sender, EventArgs e)
        {
            int fValue = int.TryParse(txtBxCurFreq.Text, out var tmp) ? tmp : 0;
            //if (comm.ComPortIsOpen) comm.WriteData(CurrentFreqCmd + fValue.ToString());
            if (!comm.WriteData(CurrentFreqCmd + fValue.ToString()))
            {
                this.ActiveControl = null;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!comm.ComPortIsOpen)
            {
                MessageBox.Show("Сначала откройте COM-порт!", "Ошибка",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // если уже идёт эксперимент — игнорируем повторный старт
            if (WrkState != WS.wsStoped)
                return;

            // Проверка связи
            if (!comm.WriteData("S"))
            {
                MessageBox.Show("Нет связи с устройством!\nПроверьте подключение и повторите попытку.",
                               "Ошибка связи",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
                return;
            }

            WrkState = WS.wsStoped;
            Thread.Sleep(5);

            Logger.Flush();
            GetControlsStates();

            Brightness = Cfg.Brightness;
            IteratinsCount = Cfg.IteratinsCount;
            NeedRandomDigits = Cfg.NeedRandomDigits;

            PrepareExcludedDigits(Cfg.DigitsForExclusion);
            InitRandom();

            // Общий стартовый тайм для логов/таймера
            QueryPerformanceCounter(out StartPc);
            LastPc = StartPc;
            LastPc2 = StartPc;

            // старт времени сессии
            _sessionRunning = true;
            _ui.Post(_ => txtBxExpTime.Text = "00:00:00.000", null);

            // Логгер — новый файл на каждый старт
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "exp_log.csv");
            Logger.Init(logPath, StartPc, Pf, false);

            // Общие команды в устройство
            comm.WriteData(BrightnessCmd + Brightness);

            // Очистки буферов
            _digitBuf.Clear();
            lock (_lock)
            {
                _freqAtKeypress.Clear();
                _pcAtKeypress.Clear();
                _isCorrect.Clear();
                _allThresholds.Clear();
            }

            // UI (кнопки)
            _ui.Post(_ =>
            {
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                lstBxStimWorkSequence.Enabled = false;
                this.ActiveControl = label6;
            }, null);

            if (Cfg.Mode == ExperimentMode.Sequential)
            {
                // Sequential: никаких Shepard звуков
                if (_SndPlayer != null) _SndPlayer.StopAll();

                Brightness = (int)numUpDwnBrightnessSequentialMode.Value;
                Cfg.Brightness = Brightness;
                comm.WriteData(BrightnessCmd + Brightness);

                Logger.WriteLog(StartPc, "session_start", "mode=Sequential");

                StartSequentialMode();
                return;
            }

            // ===== Staircase старт =====

            // стартовая цифра (цель) для лестницы
            SendRandomDigits();

            int startFreq = Cfg.StartingFreq;

            // Команды в устройство
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
            comm.WriteData(CurrentFreqCmd + startFreq);

            // Локальные переменные лестницы
            CurrentIteration = 0;
            CurrentThreshold = startFreq;
            CurrentFreq = startFreq;
            AvgThreshold = 0;

            StepTimeInTicks = (long)Pf * (long)Cfg.TimeStepFreqChngMs / 1000L;

            Logger.WriteLog(StartPc, "session_start",
                $"startFreq={startFreq};stepHz={Cfg.StepJumpFreq};stepTime={Cfg.TimeStepFreqChngMs};" +
                $"target={TwoDigitNumber};minIters={IteratinsCount};sigmaMax={Cfg.StdDevThreshold:0.###};randDigits={NeedRandomDigits}");

            _sweepDir = -1; // начнём с понижения

            if (_SndPlayer != null)
            {
                _SndPlayer.Stop("shep_up");
                _SndPlayer.Play("shep_down", true);
            }

            _ui.Post(_ =>
            {
                txtBxCurItertn.Text = "0";
                txtBxCurThresshld.Text = CurrentThreshold.ToString("0.###");
                txtBxAvrgThresshld.Text = "";
                txtBxCurSigma.Text = "";
                txtBxCurFreq.Text = CurrentFreq.ToString();
            }, null);

            WrkState = WS.wsFreqSearch;
        }

        private void StartSequentialMode()
        {
            if (Cfg.Stimulus == null || Cfg.Stimulus.Count == 0)
            {
                MessageBox.Show("Список стимулов пуст.", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

                _ui.Post(_ => 
                { 
                    btnStart.Enabled = true;
                    btnStop.Enabled = false;
                    lstBxStimWorkSequence.Enabled = true;
                }, null);
                return;
            }

            _seqIndex = 0;
            _digitBuf.Clear();

            _seqInputStarted = false;
            _seqFirstKeyPc = 0;

            _seqState = TPresentationState.Presenting;

            var s = Cfg.Stimulus[_seqIndex];

            QueryPerformanceCounter(out _seqStimulusStartPc);

            comm.WriteData(BrightnessCmd + Cfg.Brightness);
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(s.Digits));
            comm.WriteData(CurrentFreqCmd + s.Frequency);
            CurrentFreq = s.Frequency;

            // UI — номер текущего
            _ui.Post(_ =>
            {
                txtBxCurItertn.Text = "1";
                lstBxStimWorkSequence.SelectedIndex = -1;
                lstBxStimWorkSequence.SelectedIndex = 0;
            }, null);

            Logger.WriteLog(_seqStimulusStartPc,
                "stimulus_start",
                $"idx={_seqIndex};digits={s.Digits};freq={s.Frequency};dur={s.Duration};needReaction={(s.NeedReaction ? 1 : 0)}");
        }
        private void StopSequential(long pcNow, string reason, bool isManualStop)
        {
            // защита от повторного стопа (иначе будут дубль-логи)
            if (_seqState == TPresentationState.Stopped)
                return;

            _sessionRunning = false;
            _seqState = TPresentationState.Stopped;

            // на стопе частота 196, цифра остаётся текущая
            comm.WriteData(CurrentFreqCmd + "196");
            CurrentFreq = 196;

            // ЛОГИ:
            // - sequence_end только если действительно дошли до конца списка
            bool completed = (Cfg.Stimulus != null && _seqIndex >= Cfg.Stimulus.Count);

            if (completed && !isManualStop)
            {
                Logger.WriteLog(pcNow, "sequence_end", "");
            }
            else
            {
                // всё остальное — отдельное событие
                string modeReason = $"mode=Sequential;reason={reason}";
                Logger.WriteLog(pcNow, "sequence_stop", modeReason);
            }

            Logger.Flush();

            _ui.Post(_ =>
            {
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                lstBxStimWorkSequence.Enabled = true;
            }, null);
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!comm.ComPortIsOpen)
                return;

            bool seqRunning = (Cfg.Mode == ExperimentMode.Sequential) &&
                              (_seqState != TPresentationState.None) &&
                              (_seqState != TPresentationState.Stopped);

            if (Cfg.Mode == ExperimentMode.Sequential)
            {
                if (!seqRunning) return;

                QueryPerformanceCounter(out long pcNow);
                StopSequential(pcNow, "manual", true);
                return;
            }

            // ===== Staircase как было (у тебя можно оставить старый код) =====
            if (WrkState == WS.wsStoped)
                return;

            WrkState = WS.wsStoped;
            _sessionRunning = false;

            QueryPerformanceCounter(out long pcNow2);
            Logger.WriteLog(pcNow2, "session_stop", $"reason=manual;freq={CurrentFreq} target_final= {TwoDigitNumber}");

            comm.WriteData(CurrentFreqCmd + Cfg.StartingFreq.ToString());
            Logger.Flush();

            if (_SndPlayer != null)
                _SndPlayer.StopAll();

            _ui.Post(_ =>
            {
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                lstBxStimWorkSequence.Enabled = true;
            }, null);
        }

        private void btnDebug_Click(object sender, EventArgs e)
        {
            Brightness = 0;

            InitRandom();
            TwoDigitNumber = _rnd.Next(10, 100);// целевая цифра 

            int startFreq = 10;//Cfg.StartingFreq;

            // Команды в устройство
/*            comm.WriteData(BrightnessCmd + Brightness);
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
            comm.WriteData(CurrentFreqCmd + startFreq);*/


            CurrentFreq = 30;
           //
           //comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());
            //Thread.Sleep(10);
            TwoDigitNumber = _rnd.Next(10, 100);
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
            //comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());
            //проблема в том, что устанавливаемую цифру видно! Надо ковырять ардиуно-код
        }

        private void frmMain_KeyDown(object sender, KeyEventArgs e)
        {
            bool seqRunning = (Cfg.Mode == ExperimentMode.Sequential) &&
                              (_seqState != TPresentationState.None) &&
                              (_seqState != TPresentationState.Stopped);

            if (e.KeyCode == Keys.Enter && (WrkState != WS.wsStoped || seqRunning))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void cmdSend_Click(object sender, EventArgs e)
        {
            // Если порт не открыт – ничего не шлём
            if (!comm.ComPortIsOpen)
            {
                MessageBox.Show("COM-порт не открыт.", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Берём текст из поля
            var msg = txtSend.Text;
            if (string.IsNullOrWhiteSpace(msg))
                return; // Пустую команду не шлём

            // Отправка (WriteData сам добавит \n / ; при необходимости)
            comm.WriteData(msg);

            // Для удобства: выделим текст, чтобы можно было сразу перезатереть
            txtSend.SelectAll();
            txtSend.Focus();
        }

        private void btnAddStim_Click(object sender, EventArgs e)
        {
            if (Cfg.Stimulus == null)
                Cfg.Stimulus = new List<Stimulus>();

            Stimulus stm = new Stimulus(true)
            {
                Digits = GenerateAllowedTwoDigitNumber(),
                Frequency = (int)numUpDwnStartingFreqSequentialMode.Value,      // 
                Duration = (int)numUpDwnTimeStepFreqChngSequentialMode.Value,   // 
                NeedReaction = true
            };

            Cfg.Stimulus.Add(stm);

            RefreshStimulusList();

            // выделим добавленный
            lstBxStimWorkSequence.SelectedIndex = -1;
            if (lstBxStimWorkSequence.Items.Count > 0)
                lstBxStimWorkSequence.SelectedIndex = lstBxStimWorkSequence.Items.Count - 1;
        }

        private void btnRemoveStim_Click(object sender, EventArgs e)
        {
            if (Cfg.Stimulus == null || Cfg.Stimulus.Count == 0)
                return;

            if (lstBxStimWorkSequence.SelectedIndices == null || lstBxStimWorkSequence.SelectedIndices.Count == 0)
                return;

            // удаляем с конца, чтобы индексы не съезжали
            var toRemove = lstBxStimWorkSequence.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();

            foreach (int idx in toRemove)
            {
                if (idx >= 0 && idx < Cfg.Stimulus.Count)
                    Cfg.Stimulus.RemoveAt(idx);
            }

            RefreshStimulusList();

            // аккуратно восстановим выделение
            if (Cfg.Stimulus.Count > 0)
            {
                int newIdx = Math.Min(toRemove.Min(), Cfg.Stimulus.Count - 1);
                lstBxStimWorkSequence.SelectedIndex = newIdx;
            }
        }

        private void lstBxStimWorkSequence_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstBxStimWorkSequence.SelectedIndex != -1)
            {
                Stimulus stm = Cfg.Stimulus[lstBxStimWorkSequence.SelectedIndex];
                GVars.LockCtrl = true;
                numUpDwnStimulusDigit.Value = stm.Digits;
                numUpDwnStartingFreqSequentialMode.Value = stm.Frequency;
                numUpDwnTimeStepFreqChngSequentialMode.Value = stm.Duration;
                numUpDwnTrialCount.Value = lstBxStimWorkSequence.Items.Count;

                if (comm.ComPortIsOpen)
                {
                    int d = (int)numUpDwnStimulusDigit.Value;
                    comm.WriteData(CurrentDigitCmd + ToHardwareDigit(d));

                    int f = (int)numUpDwnStartingFreqSequentialMode.Value;
                    comm.WriteData(CurrentFreqCmd + ToHardwareDigit(f));
                }

                GVars.LockCtrl = false;
            }
        }

        private void btnClearStim_Click(object sender, EventArgs e)
        {
            if (Cfg.Stimulus == null)
                Cfg.Stimulus = new List<Stimulus>();
            else
                Cfg.Stimulus.Clear();

            RefreshStimulusList();
        }

        private void numUpDwnStartingFreqSequentialMode_ValueChanged(object sender, EventArgs e)
        {
            if (!GVars.LockCtrl)
            {
                bool changed = false;
                foreach (int SelectedIndex in lstBxStimWorkSequence.SelectedIndices)
                {
                    Stimulus stm = Cfg.Stimulus[SelectedIndex];
                    stm.Frequency = (int)numUpDwnStartingFreqSequentialMode.Value;
                    Cfg.Stimulus[SelectedIndex] = stm;
                    changed = true;
                }
                if (changed)
                    RefreshStimulusList();

                if (comm.ComPortIsOpen)
                {
                    int f = (int)numUpDwnStartingFreqSequentialMode.Value;
                    comm.WriteData(CurrentFreqCmd + ToHardwareDigit(f));
                }
            }
        }

        private void numUpDwnTimeStepFreqChngSequentialMode_ValueChanged(object sender, EventArgs e)
        {
            if (GVars.LockCtrl)
                return;

            if (Cfg.Stimulus == null || Cfg.Stimulus.Count == 0)
                return;

            bool changed = false;

            foreach (int idx in lstBxStimWorkSequence.SelectedIndices)
            {
                if (idx < 0 || idx >= Cfg.Stimulus.Count)
                    continue;

                Stimulus stm = Cfg.Stimulus[idx];
                stm.Duration = (int)numUpDwnTimeStepFreqChngSequentialMode.Value;
                Cfg.Stimulus[idx] = stm;

                changed = true;
            }

            if (changed)
                RefreshStimulusList();
        }

        private bool TryParseStimulusLine(string line, out Stimulus stim)
        {
            stim = new Stimulus(true);

            if (string.IsNullOrWhiteSpace(line))
                return false;

            var parts = line.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var p in parts)
            {
                var kv = p.Split('=');
                if (kv.Length != 2)
                    continue;

                string key = kv[0].Trim().ToLower();
                string val = kv[1].Trim();

                switch (key)
                {
                    case "digits":
                        if (int.TryParse(val, out int d))
                            stim.Digits = d;
                        break;

                    case "freq":
                        if (int.TryParse(val, out int f))
                            stim.Frequency = f;
                        break;

                    case "dur":
                        if (int.TryParse(val, out int dur))
                            stim.Duration = dur;
                        break;

                    case "needreaction":
                        stim.NeedReaction = val.Equals("true", StringComparison.OrdinalIgnoreCase);
                        break;
                }
            }

            return true;
        }

        private void RefreshStimulusListBox()
        {
            lstBxStimWorkSequence.Items.Clear();

            for (int i = 0; i < Cfg.Stimulus.Count; i++)
            {
                var s = Cfg.Stimulus[i];
                lstBxStimWorkSequence.Items.Add($"{i + 1}. {s.Digits}");
            }
        }

        private void btnLoadListStimulus_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Stimulus list (*.txt)|*.txt|All files (*.*)|*.*";
                ofd.Title = "Загрузить список стимулов";

                if (ofd.ShowDialog(this) != DialogResult.OK)
                    return;

                string[] rawLines = File.ReadAllLines(ofd.FileName, Encoding.UTF8);

                // убираем пустые строки
                var lines = rawLines
                    .Select(x => (x ?? "").Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                if (lines.Count == 0)
                {
                    MessageBox.Show("Файл пуст.", "Загрузка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Первая строка — шапка
                // Ожидаем: Digits;Frequency;Duration;NeedReaction
                // Если там что-то другое — не умираем, но предупредим.
                string header = lines[0];
                bool headerOk = header.Replace(" ", "").Equals("Digits;Frequency;Duration;NeedReaction", StringComparison.OrdinalIgnoreCase);

                var newList = new List<Stimulus>();
                int bad = 0;

                // начинаем со 2-й строки
                for (int i = 1; i < lines.Count; i++)
                {
                    string line = lines[i];

                    var parts = line.Split(new[] { ';' }, StringSplitOptions.None);
                    if (parts.Length < 4) { bad++; continue; }

                    if (!int.TryParse(parts[0].Trim(), out int digits)) { bad++; continue; }
                    if (!int.TryParse(parts[1].Trim(), out int freq)) { bad++; continue; }
                    if (!int.TryParse(parts[2].Trim(), out int dur)) { bad++; continue; }

                    string needStr = parts[3].Trim();
                    bool need;
                    if (needStr == "1") need = true;
                    else if (needStr == "0") need = false;
                    else if (needStr.Equals("true", StringComparison.OrdinalIgnoreCase)) need = true;
                    else if (needStr.Equals("false", StringComparison.OrdinalIgnoreCase)) need = false;
                    else { bad++; continue; }

                    // минимальная гигиена
                    if (digits < 0 || digits > 99) { bad++; continue; }
                    if (freq < 0) { bad++; continue; }
                    if (dur < 0) { bad++; continue; }

                    var s = new Stimulus(true)
                    {
                        Digits = digits,
                        Frequency = freq,
                        Duration = dur,
                        NeedReaction = need
                    };

                    newList.Add(s);
                }

                if (newList.Count == 0)
                {
                    MessageBox.Show("Не удалось загрузить ни одного стимула.\nПроверь формат строк.",
                        "Загрузка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Истина теперь тут
                Cfg.Stimulus = newList;

                // UI — только витрина
                RefreshStimulusList();

                if (lstBxStimWorkSequence.Items.Count > 0)
                    lstBxStimWorkSequence.SelectedIndex = 0;

                // Сообщение по итогам
                if (!headerOk || bad > 0)
                {
                    var msg = $"Загружено стимулов: {newList.Count}";
                    if (bad > 0) msg += $"\nПропущено строк: {bad}";
                    if (!headerOk) msg += "\n\nВнимание: шапка файла отличается от ожидаемой (Digits;Frequency;Duration;NeedReaction).";

                    MessageBox.Show(msg, "Загрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        public string StimulusToLine(Stimulus s)
        {
            return
                $"digits={s.Digits};freq={s.Frequency};dur={s.Duration};needReaction={(s.NeedReaction ? "true" : "false")}";
        }
        private void btnSaveListStimulus_Click(object sender, EventArgs e)
        {
            if (Cfg.Stimulus == null || Cfg.Stimulus.Count == 0)
            {
                MessageBox.Show("Список стимулов пуст.", "Сохранение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Stimulus list (*.txt)|*.txt|All files (*.*)|*.*";
                sfd.Title = "Сохранить список стимулов";
                sfd.FileName = "stimulus_list.txt";

                if (sfd.ShowDialog(this) != DialogResult.OK)
                    return;

                var lines = new List<string>();

                // ШАПКА (как просил): обозначения полей
                lines.Add("Digits;Frequency;Duration;NeedReaction");

                // ДАННЫЕ: только значения
                foreach (var s in Cfg.Stimulus)
                {
                    int need = s.NeedReaction ? 1 : 0;
                    lines.Add($"{s.Digits};{s.Frequency};{s.Duration};{need}");
                }

                File.WriteAllLines(sfd.FileName, lines, Encoding.UTF8);
            }
        }

        private void numUpDwnStimulusDigit_ValueChanged(object sender, EventArgs e)
        {
            if (!GVars.LockCtrl)
            {
                bool changed = false;
                foreach (int SelectedIndex in lstBxStimWorkSequence.SelectedIndices)
                {
                    Stimulus stm = Cfg.Stimulus[SelectedIndex];
                    stm.Digits = (int)numUpDwnStimulusDigit.Value;
                    Cfg.Stimulus[SelectedIndex] = stm;
                    changed = true;
                }
                if (changed)
                    RefreshStimulusList();

                if (comm.ComPortIsOpen)
                {
                    int d = (int)numUpDwnStimulusDigit.Value;
                    comm.WriteData(CurrentDigitCmd + ToHardwareDigit(d));
                }
            }
        }

        private void numUpDwnBrightnessSequentialMode_ValueChanged(object sender, EventArgs e)
        {
            if (GVars.LockCtrl)
                return;

            int b = (int)numUpDwnBrightnessSequentialMode.Value;

            Cfg.Brightness = b;
            Brightness = b;

            //срузу применяем, чтобы видеть
            if (comm.ComPortIsOpen)
            {
                comm.WriteData(BrightnessCmd + b);
            }
        }
        private void lstBxStimWorkSequence_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
                return;

            ListBox lb = (ListBox)sender;

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // Задаем цвета для активного и неактивного ListBox
            Color back;
            Color fore;

            if (lb.Enabled) // Если ListBox активен
            {
                if (selected)
                {
                    // Бледно-синий для выбранной строки
                    back = Color.FromArgb(180, 200, 230);
                    fore = Color.Black;
                }
                else
                {
                    back = lb.BackColor;
                    fore = lb.ForeColor;
                }
            }
            else // Если ListBox не активен
            {
                if (selected)
                {
                    back = Color.FromArgb(180, 200, 230);  // Бледно-синий для выделенной строки
                    fore = Color.Gray;  // Для неактивного выделения серый цвет
                }
                else
                {
                    back = lb.BackColor;
                    fore = Color.Gray; // Темно-серый текст для неактивного состояния
                }
            }

            using (SolidBrush b = new SolidBrush(back))
                e.Graphics.FillRectangle(b, e.Bounds);

            string text = lb.Items[e.Index].ToString();

            TextRenderer.DrawText(
                e.Graphics,
                text,
                lb.Font,
                e.Bounds,
                fore,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

            e.DrawFocusRectangle();
        }
    }
}
