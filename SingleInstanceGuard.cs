using System.Threading;
using System.Windows.Forms;

namespace DFlasher
{
    internal static class SingleInstanceGuard
    {
        private static Mutex _mutex;

        /// <summary>
        /// Проверка, что приложение ещё не запущено.
        /// Возвращает true, если это первый экземпляр.
        /// Возвращает false, если экземпляр уже есть (даже зависший).
        /// </summary>
        internal static bool EnsureSingleInstance()
        {
            bool createdNew = false;

            // УНИКАЛЬНОЕ имя мьютекса во всей системе
            _mutex = new Mutex(
                true,                          // сразу пытаемся захватить
                "Global\\DFlasher_SingleInstanceMutex",
                out createdNew);

            if (!createdNew)
            {
                // Уже есть процесс, который держит мьютекс
                MessageBox.Show(
                    "Программа DFlasher уже запущена и находится в памяти.",
                    "DFlasher",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }

        /// <summary>
        /// Освобождение мьютекса при завершении работы.
        /// </summary>
        internal static void Release()
        {
            if (_mutex != null)
            {
                try
                {
                    _mutex.ReleaseMutex();
                    _mutex.Dispose();
                }
                catch
                {
                    // игнорируем ошибки при завершении
                }
                _mutex = null;
            }
        }
    }
}
