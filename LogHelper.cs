using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace DFlasher
{
    internal static class Logger
    {
        private static BlockingCollection<string> _queue = new BlockingCollection<string>(new ConcurrentQueue<string>());
        private static Task _task;
        private static string _fileName;
        private static bool _started = false;
        private static readonly object _initLock = new object();

        // Для конвертации QPC→мс относительно старта
        private static long _pf;       // QueryPerformanceFrequency
        private static long _startPc;  // QueryPerformanceCounter на старте

        /// <summary>
        /// Инициализация логгера: файл, стартовый тик и частота счетчика.
        /// Вызывать один раз при старте эксперимента.
        /// </summary>
        public static void Init(string fileName, long startPc, long pf, bool append = true)
        {
            lock (_initLock)
            {
                if (_started)
                    return;

                _startPc = startPc;
                _pf = pf;

                // Разобрать путь/имя/расширение
                string dir = Path.GetDirectoryName(fileName);
                if (string.IsNullOrEmpty(dir))
                    dir = AppDomain.CurrentDomain.BaseDirectory;

                string baseName = Path.GetFileNameWithoutExtension(fileName);
                if (string.IsNullOrEmpty(baseName))
                    baseName = "exp_log";

                string ext = Path.GetExtension(fileName);
                if (string.IsNullOrEmpty(ext))
                    ext = ".csv";

                Directory.CreateDirectory(dir);

                // Если append==false — создаём новое имя с меткой времени
                if (!append)
                {
                    string ts = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                    _fileName = Path.Combine(dir, baseName + "_" + ts + ext);
                }
                else
                {
                    _fileName = Path.Combine(dir, baseName + ext);
                }

                bool writeHeader = true;
                if (append && File.Exists(_fileName))
                {
                    // При дозаписи шапку пишем только если файл пуст
                    var fi = new FileInfo(_fileName);
                    writeHeader = fi.Length == 0;
                }

                _task = Task.Factory.StartNew(delegate
                {
                    using (var sw = new StreamWriter(_fileName, append, Encoding.UTF8))
                    {
                        sw.AutoFlush = true;

                        if (writeHeader)
                            sw.WriteLine("timestamp_ms;event;details");

                        foreach (var line in _queue.GetConsumingEnumerable())
                            sw.WriteLine(line);
                    }
                }, TaskCreationOptions.LongRunning);

                _started = true;
            }
        }

        /// <summary>
        /// Завершить логгер и дождаться записи очереди.
        /// </summary>
        public static void Flush()
        {
            if (!_started) return;
            _queue.CompleteAdding();
            try { _task.Wait(); } catch { /* игнорируем */ }
            _started = false;
        }

        // ================= ПУБЛИЧНЫЕ ЗАПИСИ =================

        /// <summary>
        /// Универсальная запись: details.
        /// pcNow — текущее значение QueryPerformanceCounter (фактическое время события!).
        /// </summary>
        public static void WriteLog(long pcNow, string eventName, string details)
        {
            Enqueue(Format(pcNow, eventName, details));
        }

        // Перегрузки «как у вас», но без DateTime.Now — время берём из pcNow/QPC.
        public static void WriteLog(long pcNow, string action, int errorCode, string errorDescription)
        {
            string details = "action=" + action + ";code=" + errorCode.ToString() + ";desc=" + errorDescription;
            Enqueue(Format(pcNow, "event", details));
        }

        public static void WriteLog(long pcNow, string action, float code1, float code2, string description)
        {
            string details = "action=" + action + ";code1=" + code1.ToString() + ";code2=" + code2.ToString() + ";desc=" + description;
            Enqueue(Format(pcNow, "event", details));
        }

        public static void WriteLog(long pcNow, string action, long code, string description)
        {
            string details = "action=" + action + ";code=" + code.ToString() + ";desc=" + description;
            Enqueue(Format(pcNow, "event", details));
        }

        // ================= ВСПОМОГАТЕЛЬНОЕ =================

        private static void Enqueue(string line)
        {
            if (!_started) return; // можно и бросать исключение, но безопаснее молча игнорить до Init
            if (!_queue.IsAddingCompleted)
                _queue.Add(line);
        }

        private static string Format(long pcNow, string ev, string details)
        {
            double ms = PcToMs(pcNow);
            // CSV: миллисекунды с точностью до тысячных
            return ms.ToString("0.###") + ";" + ev + ";" + details;
        }

        private static double PcToMs(long pcNow)
        {
            if (_pf == 0) return 0.0;
            long delta = pcNow - _startPc;
            return ((double)delta) * 1000.0 / (double)_pf;
        }
    }
}
