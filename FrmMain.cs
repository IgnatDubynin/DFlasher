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

        public enum WS { wsStoped, wsFreqSearch, wsWaiting, wsFreqFound, wsFreqSearchByStep, wsReset };
        public WS WrkState = WS.wsStoped;

        public long Pf;
        public long Pc;
        public long LastPc;
        public long LastPc2;
        public long StartPc;

        private SynchronizationContext _ui;

        private long _intervalTicks; // период обновления для отображения времени
        long StepTimeInTicks;

        private int CurrentFreq;
        private int CurrentIteration;
        private int CurrentThreshold;
        private static int NeedSwapDigits;
        private double AvgThreshold = 0;
        private List<double> _allThresholds = new List<double>();
        // направление «ползания»: -1 — вниз, +1 — вверх
        private int _sweepDir = -1;          // начнём с понижения
        //private bool _jumpUpNext = true;     // первый скачок — вверх
        private Random _rnd;

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

            _ui = SynchronizationContext.Current;
            _intervalTicks = Pf / 100; // 10 мс

            QueryPerformanceFrequency(out Pf);
            QueryPerformanceCounter(out LastPc);

            timeBeginPeriod(1);
            mHandler = new TimerEventHandler(TmrCallbckForMainProcessing);
            mTimerId = timeSetEvent(1, 1, mHandler, IntPtr.Zero, EVENT_TYPE);

            _hookProc = KeyboardHookProc;
            var hMod = GetModuleHandle(null);
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, hMod, 0);

            InitializeComponent();

            this.KeyPreview = true;

        }
        private void TmrCallbckForMainProcessing(int id, int msg, IntPtr user, int dw1, int dw2)
        {
            if (WrkState != WS.wsStoped)
            {
                QueryPerformanceCounter(out Pc);

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
                                    // Не удалось отправить - останавливаем эксперимент
                                    SafeStopExperiment("Потеря связи с устройством");
                                    return;
                                }
                                CurrentFreq--;
                                // новая случайная цифра
                                TwoDigitNumber = _rnd.Next(10, 100);
                                comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
                            }
                            else
                            {
                                if (_SndPlayer != null)
                                {
                                    _SndPlayer.StopAll();
                                }
                                // упёрлись в минимум — развернуть направление
                                // _sweepDir = +1;
                                //WrkState = WS.wsStoped;
                            }
                        }
                        else // _sweepDir > 0
                        {
                            if (CurrentFreq < Cfg.StartingFreq)
                            {
                                if (!comm.WriteData(IncCurrentFreqCmd))
                                {
                                    SafeStopExperiment("Потеря связи с устройством");
                                    return;
                                }
                                CurrentFreq++;
                                // новая случайная цифра
                                TwoDigitNumber = _rnd.Next(10, 100);
                                comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
                            }
                            else
                            {
                                if (_SndPlayer != null)
                                {
                                    _SndPlayer.StopAll();
                                }
                                // упёрлись в максимум — развернуть направление
                                //_sweepDir = -1;
                                //WrkState = WS.wsStoped;
                            }
                        }

                        LastPc2 = Pc;
                    }
                }

                if (Pc >= LastPc + _intervalTicks)
                {
                    double elapsedSeconds = (double)(Pc - StartPc) / Pf;
                    string formatted = TimeSpan.FromSeconds(elapsedSeconds).ToString(@"hh\:mm\:ss\.fff");

                    _ui.Post(_ => txtBxExpTime.Text = formatted, null);

                    LastPc = Pc;
                }
            }
        }
        private void SafeStopExperiment(string reason)
        {
            // Сохраняем состояние
            bool wasRunning = (WrkState != WS.wsStoped);

            // Останавливаем эксперимент
            WrkState = WS.wsStoped;

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
        // ==== Глобальный хук: ловим любую клавишу, цифры и Enter ====
        private IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            const int WM_KEYDOWN = 0x0100;
            if (wParam != (IntPtr)WM_KEYDOWN)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            if (!comm.ComPortIsOpen || WrkState == WS.wsStoped)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            KBDLLHOOKSTRUCT k = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            int vk = k.vkCode;

            QueryPerformanceCounter(out long pcNow);
            double msFromStart = (pcNow - StartPc) * 1000.0 / Pf;
            int freqNow = CurrentFreq;

            // Первый "любой" нажим — переводим в wsWaiting и фиксируем событие
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
                    {
                        // добавляем именно символ цифры, а не (char)vk
                        _digitBuf.Append((char)('0' + digit));
                    }
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

                            // ==== итерация ====
                            CurrentIteration++;

                            // текущий порог = текущая частота (на момент первого тычка)
                            CurrentThreshold = freqNow;

                            // средний порог
                            _allThresholds.Add(CurrentThreshold);
                            AvgThreshold = _allThresholds.Average();

                            double sigma = ComputeSigmaHz();

                            // вывести на контролы
                            _ui.Post(_ =>
                            {
                                txtBxCurItertn.Text = CurrentIteration.ToString();
                                txtBxCurThresshld.Text = CurrentThreshold.ToString("0.###");
                                txtBxAvrgThresshld.Text = AvgThreshold.ToString("0.###");
                                txtBxCurSigma.Text = sigma.ToString("0.###");
                            }, null);

                            int minIters = 3;

                            // 0 тоже «хорошо», но нужен минимум итераций
                            if (CurrentIteration >= minIters && sigma <= Cfg.StdDevIn3Itrtn)
                            {
                                // ЛОГ ОТВЕТА БЕЗ ПРЫЖКА
                                string details = string.Format(
                                    "input={0};target={1};correct=1;CurFreq={2};σ={3:0.###}",
                                    entered, TwoDigitNumber, freqNow, sigma);

                                Logger.WriteLog(pcNow, "answer", details);

                                WrkState = WS.wsStoped;
                                Logger.WriteLog(pcNow, "stop",
                                    string.Format("σ-threshold reached;freq={0};σ={1:0.###}", freqNow, sigma));

                                // Вернуть стимуляцию к начальному значению (если так задумано)
                                comm.WriteData(CurrentFreqCmd + Cfg.StartingFreq.ToString());

                                Logger.Flush();

                                if (_SndPlayer != null)
                                {
                                    // на всякий случай глушим ползущие Shepard'ы
                                    _SndPlayer.Stop("shep_up");
                                    _SndPlayer.Stop("shep_down");

                                    // наш жирный органный пабамм
                                    _SndPlayer.Play("success_pabam", false); // false — без зацикливания
                                }

                                // UI: обновим поля
                                _ui.Post(_ =>
                                {
                                    //txtBxCurItertn.Text = CurrentIteration.ToString();
                                    //txtBxCurThresshld.Text = CurrentFreq.ToString("0.###");
                                    //txtBxCurSigma.Text = "";

                                    // кнопки в состояние «можно снова стартовать»
                                    btnStart.Enabled = true;
                                    btnStop.Enabled = false;
                                }, null);
                            }
                            else
                            {
                                // ЛОГ ОТВЕТА: только текущая частота и σ
                                string details = string.Format(
                                    "input={0};target={1};correct=1;CurFreq={2};σ={3:0.###}",
                                    entered, TwoDigitNumber, freqNow, sigma);

                                Logger.WriteLog(pcNow, "answer", details);

                                // текущий знак ползания
                                int oldDir = _sweepDir;

                                // частота для следующего цикла (новый старт после прыжка)
                                int newStartFreq = CurrentFreq + oldDir * Cfg.StepJumpFreq;

                                // ОГРАНИЧЕНИЕ: [1; Cfg.StartingFreq]
                                if (newStartFreq < 1)
                                    newStartFreq = 1;
                                if (newStartFreq > Cfg.StartingFreq)
                                    newStartFreq = Cfg.StartingFreq;

                                // выполняем сам прыжок
                                CurrentFreq = newStartFreq;
                                comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());

                                // новая случайная цифра
                                TwoDigitNumber = _rnd.Next(10, 100);
                                comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));

                                // новое направление ползания
                                _sweepDir = -oldDir;
                                string workDir = (_sweepDir < 0 ? "down" : "up");


                                if (_SndPlayer != null)
                                {
                                    if (_sweepDir < 0)
                                    {
                                        // вниз: включаем down, глушим up
                                        _SndPlayer.Stop("shep_up");
                                        _SndPlayer.Play("shep_down", true);
                                    }
                                    else if (_sweepDir > 0)
                                    {
                                        // вверх: включаем up, глушим down
                                        _SndPlayer.Stop("shep_down");
                                        _SndPlayer.Play("shep_up", true);
                                    }
                                }

                                // лог прыжка и направления работы
                                Logger.WriteLog(pcNow, "jump",
                                    string.Format("NewStartFreq={0};dir={1}", CurrentFreq, workDir));

                                LastPc2 = pcNow;
                                WrkState = WS.wsFreqSearch;
                            }
                        }
                        else // неверная цифра — всё равно делаем скачок, но без учёта в пороге/σ
                        {
                            // лог ответа: без σ, но с текущей частотой
                            string details = string.Format(
                                "input={0};target={1};correct=0;CurFreq={2}",
                                entered, TwoDigitNumber, freqNow);

                            Logger.WriteLog(pcNow, "answer", details);

                            int oldDir = _sweepDir;

                            // частота для следующего цикла (новый старт после прыжка)
                            int newStartFreq = CurrentFreq + oldDir * Cfg.StepJumpFreq;

                            // ОГРАНИЧЕНИЕ: [1; Cfg.StartingFreq]
                            if (newStartFreq < 1)
                                newStartFreq = 1;
                            if (newStartFreq > Cfg.StartingFreq)
                                newStartFreq = Cfg.StartingFreq;

                            CurrentFreq = newStartFreq;
                            comm.WriteData(CurrentFreqCmd + CurrentFreq.ToString());

                            TwoDigitNumber = _rnd.Next(10, 100);
                            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));

                            _sweepDir = -oldDir;
                            string workDir = (_sweepDir < 0 ? "down" : "up");

                            if (_SndPlayer != null)
                            {
                                if (_sweepDir < 0)
                                {
                                    // вниз: включаем down, глушим up
                                    _SndPlayer.Stop("shep_up");
                                    _SndPlayer.Play("shep_down", true);
                                }
                                else if (_sweepDir > 0)
                                {
                                    // вверх: включаем up, глушим down
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
                        // Enter на неполных 2 цифрах — игнор
                    }
                }
                // прочие клавиши игнорируем
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
        private void SetControlsStates()
        {
            rdoHex.Checked = true;
            // rdoText.Checked = true;
            cmdSend.Enabled = false;
            cmdClose.Enabled = false;

            numUpDwnStartingFreq.Value = Cfg.StartingFreq;
            numUpDwnStepFreqHz.Value = Cfg.StepJumpFreq;
            numUpDwnTimeStepFreqChng.Value = Cfg.TimeStepFreqChngMs;
            numUpDwnStepClarifyThrshld.Value = Cfg.StepClarifyThrshldHz;
            numUpDwnStdDevIn3Itrtn.Value = Cfg.StdDevIn3Itrtn;
            numUpDwnBrightness.Value = Cfg.Brightness;

            // Режим эксперимента
            rdoBtnStaircase.Checked = (Cfg.Mode == ExperimentMode.Staircase);
            rdoBtnSequential.Checked = (Cfg.Mode == ExperimentMode.Sequential);

            // Sequential настройки (только нужные)
            numUpDwnStartingFreqSequentialMode.Value = Cfg.SequentialStartingFreq;
            numUpDwnTrialCount.Value = Cfg.SequentialTrials;
            numUpDwnBrightnessSequentialMode.Value = Cfg.SequentialBrightness;
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
            Cfg.StdDevIn3Itrtn = (int)numUpDwnStdDevIn3Itrtn.Value;
            Cfg.Brightness = (int)numUpDwnBrightness.Value;
            // Sequential настройки
            Cfg.SequentialStartingFreq = (int)numUpDwnStartingFreqSequentialMode.Value;
            Cfg.SequentialTrials = (int)numUpDwnTrialCount.Value;
            Cfg.SequentialBrightness = (int)numUpDwnBrightnessSequentialMode.Value;
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

            //если уже идёт эксперимент — игнорируем повторный старт
            if (WrkState != WS.wsStoped)
                return;

            // ПРОВЕРКА 2: Связь работает (тестовая команда)
            if (!comm.WriteData("S"))  // Команда статуса
            {
                MessageBox.Show("Нет связи с устройством!\n" +
                               "Проверьте подключение и повторите попытку.",
                               "Ошибка связи",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
                return;
            }

            // Если предыдущая сессия была, закрываем лог перед новым стартом
            Logger.Flush();

            GetControlsStates();

            Brightness = Cfg.Brightness;

            InitRandom();
            TwoDigitNumber = _rnd.Next(10, 100); // целевая цифра 

            int startFreq = Cfg.StartingFreq;

            // Команды в устройство
            comm.WriteData(BrightnessCmd + Brightness);
            comm.WriteData(CurrentDigitCmd + ToHardwareDigit(TwoDigitNumber));
            comm.WriteData(CurrentFreqCmd + startFreq);

            // Локальные/глобальные переменные эксперимента
            CurrentIteration = 0;
            CurrentThreshold = startFreq;
            CurrentFreq = startFreq;   // синхронизируем переменную с фактической частотой
            AvgThreshold = 0;

            // Пересчёт интервала (long!)
            StepTimeInTicks = (long)Pf * (long)Cfg.TimeStepFreqChngMs / 1000L;

            // Стартовые тики
            QueryPerformanceCounter(out StartPc);
            LastPc = StartPc;
            LastPc2 = StartPc;         // сбрасываем "метку" для ползания

            // Логгер — новый файл на каждый старт (append = false)
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "exp_log.csv");
            Logger.Init(logPath, StartPc, Pf, false);
            Logger.WriteLog(StartPc, "session_start",
                $"startFreq={startFreq};stepHz={Cfg.StepJumpFreq};stepTime={Cfg.TimeStepFreqChngMs};target={TwoDigitNumber}");

            // <<< УБРАНО: не пишем фейковый jump в t=0, он всё только путает
            // Logger.WriteLog(StartPc, "jump",
            //     $"NewStartFreq={CurrentFreq};dir={(_sweepDir < 0 ? "down" : "up")}");

            // Очистки буферов
            _digitBuf.Clear();
            lock (_lock)
            {
                _freqAtKeypress.Clear();
                _pcAtKeypress.Clear();
                _isCorrect.Clear();
                _allThresholds.Clear();
            }

            // Направление/очередь скачков
            _sweepDir = -1; // начнём с понижения

            if (_SndPlayer != null)
            {
                _SndPlayer.Stop("shep_up");
                _SndPlayer.Play("shep_down", true);
            }

            // Обновим UI
            _ui.Post(_ =>
            {
                txtBxCurItertn.Text = "0";
                txtBxCurThresshld.Text = CurrentThreshold.ToString("0.###");
                txtBxAvrgThresshld.Text = "";
                txtBxCurSigma.Text = "";
                txtBxCurFreq.Text = CurrentFreq.ToString();

                // <<< опционально, но очень полезно:
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                this.ActiveControl = label6;
            }, null);
            // С этого момента считаем, что сессия реально запускается
            WrkState = WS.wsFreqSearch;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!comm.ComPortIsOpen)
                return;

            // Если уже остановлены — ничего не делаем
            if (WrkState == WS.wsStoped)
                return;

            WrkState = WS.wsStoped;

            // Зафиксируем текущее состояние
            QueryPerformanceCounter(out long pcNow);
 //           Logger.WriteLog(pcNow, "session_stop", $"reason=manual;freq={CurrentFreq}");
            Logger.WriteLog(pcNow, "session_stop", $"reason=manual;freq={CurrentFreq} target_final= {TwoDigitNumber}");


            // Вернуть стимуляцию к начальному значению (если так задумано)
            comm.WriteData(CurrentFreqCmd + Cfg.StartingFreq.ToString());
                        Logger.WriteLog(pcNow, "session_stop", $"reason=manual;freq={CurrentFreq} target_final= {TwoDigitNumber}");


            Logger.Flush();

            if (_SndPlayer != null)
            {
                _SndPlayer.StopAll();
            }

            // UI: обновим поля
            _ui.Post(_ =>
            {
                //txtBxCurItertn.Text = CurrentIteration.ToString();
                //txtBxCurThresshld.Text = CurrentFreq.ToString("0.###");
                //txtBxCurSigma.Text = "";

                // кнопки в состояние «можно снова стартовать»
                btnStart.Enabled = true;
                btnStop.Enabled = false;
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
            // Если идёт эксперимент, не даём Enter сработать как "клик по кнопке"
            if (e.KeyCode == Keys.Enter && WrkState != WS.wsStoped)
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
    }
}
