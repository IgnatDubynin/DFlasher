using System;
using System.IO.Ports;
using System.Threading;
using System.IO;

namespace DFlasher
{
    /// <summary>
    /// Класс для аварийной разблокировки при зависании COM-порта (пока не нужен)
    /// Предоставляет механизм принудительного восстановления работы программы
    /// </summary>
    public static class EmergencyUnlocker
    {
        private static readonly object _syncLock = new object();
        private static volatile bool _isEmergencyMode = false;
        private static volatile bool _isForceClosing = false;

        /// <summary>
        /// Флаг аварийного режима
        /// </summary>
        public static bool IsEmergencyMode => _isEmergencyMode;

        /// <summary>
        /// Войти в аварийный режим
        /// </summary>
        public static void EnterEmergencyMode()
        {
            if (_isEmergencyMode) return;

            lock (_syncLock)
            {
                if (_isEmergencyMode) return;

                _isEmergencyMode = true;
                _isForceClosing = true;

                try
                {
                    LogEmergencyEvent("Entering emergency mode");
                    ForceCloseAllSerialPorts();
                    LogEmergencyEvent("Emergency mode activated");
                }
                catch (Exception ex)
                {
                    LogEmergencyEvent($"Error in emergency mode: {ex.Message}");
                }
                finally
                {
                    _isForceClosing = false;
                }
            }
        }

        /// <summary>
        /// Выйти из аварийного режима (например, после восстановления)
        /// </summary>
        public static void ExitEmergencyMode()
        {
            lock (_syncLock)
            {
                _isEmergencyMode = false;
                LogEmergencyEvent("Exiting emergency mode");
            }
        }

        /// <summary>
        /// Принудительное закрытие всех COM-портов
        /// </summary>
        private static void ForceCloseAllSerialPorts()
        {
            string[] ports = Array.Empty<string>();

            try
            {
                ports = SerialPort.GetPortNames();
                LogEmergencyEvent($"Found {ports.Length} COM ports");
            }
            catch (Exception ex)
            {
                LogEmergencyEvent($"Failed to get port list: {ex.Message}");
                return;
            }

            foreach (var portName in ports)
            {
                if (_isForceClosing && portName.StartsWith("COM", StringComparison.OrdinalIgnoreCase))
                {
                    ForceClosePort(portName);
                }
            }
        }

        /// <summary>
        /// Принудительное закрытие конкретного порта
        /// </summary>
        private static void ForceClosePort(string portName)
        {
            SerialPort port = null;

            try
            {
                port = new SerialPort(portName)
                {
                    ReadTimeout = 100,
                    WriteTimeout = 100
                };

                if (port.IsOpen)
                {
                    LogEmergencyEvent($"Force closing port: {portName}");

                    // Сбрасываем управляющие линии
                    port.DtrEnable = false;
                    port.RtsEnable = false;
                    port.DiscardInBuffer();
                    port.DiscardOutBuffer();

                    // Пытаемся закрыть нормально
                    Thread closeThread = new Thread(() =>
                    {
                        try
                        {
                            port.Close();
                        }
                        catch { /* Игнорируем в отдельном потоке */ }
                    });

                    closeThread.Start();

                    // Даем 500мс на нормальное закрытие
                    if (!closeThread.Join(500))
                    {
                        closeThread.Interrupt();
                        LogEmergencyEvent($"Port {portName} required thread interrupt");
                    }

                    port.Dispose();
                    LogEmergencyEvent($"Port {portName} force closed");
                }
                else
                {
                    port.Dispose();
                }
            }
            catch (Exception ex)
            {
                LogEmergencyEvent($"Error closing port {portName}: {ex.GetType().Name}");

                // Последняя попытка через Dispose
                try { port?.Dispose(); } catch { }
            }
            finally
            {
                port = null;
            }
        }

        /// <summary>
        /// Логирование аварийных событий
        /// </summary>
        private static void LogEmergencyEvent(string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] EMERGENCY: {message}\n";

                // Записываем в несколько мест для надежности
                File.AppendAllText("emergency.log", logEntry);

                // Также в отдельный файл с датой
                string datedFile = $"emergency_{DateTime.Now:yyyyMMdd}.log";
                File.AppendAllText(datedFile, logEntry);

                // Консоль для отладки
                System.Diagnostics.Debug.WriteLine(logEntry);
            }
            catch
            {
                // Если даже логгирование не работает - последняя попытка
                try
                {
                    System.Diagnostics.Debug.WriteLine($"EMERGENCY LOG FAILED: {message}");
                }
                catch { }
            }
        }

        /// <summary>
        /// Проверка, находится ли программа в аварийном режиме
        /// (Thread-safe версия с дополнительной проверкой)
        /// </summary>
        public static bool SafeCheckEmergencyMode()
        {
            // Двойная проверка для производительности
            if (!_isEmergencyMode) return false;

            lock (_syncLock)
            {
                return _isEmergencyMode;
            }
        }

        /// <summary>
        /// Выполнить действие с защитой от зависания
        /// </summary>
        public static bool ExecuteWithTimeout(Action action, int timeoutMs = 5000)
        {
            if (SafeCheckEmergencyMode())
            {
                LogEmergencyEvent("Skipping action - emergency mode active");
                return false;
            }

            Thread actionThread = null;
            bool completed = false;

            try
            {
                actionThread = new Thread(() =>
                {
                    try
                    {
                        action();
                        completed = true;
                    }
                    catch (ThreadInterruptedException)
                    {
                        LogEmergencyEvent("Action interrupted by timeout");
                    }
                    catch (Exception ex)
                    {
                        LogEmergencyEvent($"Action failed: {ex.Message}");
                    }
                });

                actionThread.Start();

                if (actionThread.Join(timeoutMs))
                {
                    return completed;
                }
                else
                {
                    // Таймаут - прерываем поток
                    actionThread.Interrupt();

                    // Даем время на завершение
                    if (!actionThread.Join(1000))
                    {
                        actionThread.Abort(); // Крайняя мера
                    }

                    LogEmergencyEvent($"Action timeout after {timeoutMs}ms");
                    EnterEmergencyMode();
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogEmergencyEvent($"ExecuteWithTimeout error: {ex.Message}");
                EnterEmergencyMode();
                return false;
            }
            finally
            {
                actionThread = null;
            }
        }
    }
}
