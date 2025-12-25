using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Text.RegularExpressions;

namespace DFlasher
{
    public static class SerialHelpers
    {
        // обычный Regex, без target-typed new
        private static readonly Regex ComRx = new Regex(@"\((COM\d+)\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Возвращает список COM-портов (например, "COM8"),
        /// у которых в имени устройства есть "CH340".
        /// </summary>
        public static string[] FindCh340Ports()
        {
            var result = new List<string>();

            try
            {
                // старый стиль using (...) { }
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT Name FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"))
                {
                    foreach (ManagementObject mo in searcher.Get())
                    {
                        object nameObj = mo["Name"];
                        if (nameObj == null) continue;

                        string name = nameObj.ToString();
                        if (name.IndexOf("CH340", StringComparison.OrdinalIgnoreCase) < 0)
                            continue;

                        Match m = ComRx.Match(name);
                        if (!m.Success) continue;

                        string port = m.Groups[1].Value; // типа "COM8"
                        if (!result.Exists(p => string.Equals(p, port, StringComparison.OrdinalIgnoreCase)))
                            result.Add(port);
                    }
                }
            }
            catch
            {
                // если что-то пошло не так — вернём пустой список
            }

            return result.ToArray();
        }
        private class BtPortInfo
        {
            public string Name;
            public string PnpDeviceId;
            public int ComNumber;
        }

        /// <summary>
        /// Находит Bluetooth COM-порт HC-06.
        /// Возвращает строку вида "COM9".
        /// Ищет COM-порты, чьи PNPDeviceID начинаются с "BTHENUM\".
        /// Среди них выбирает порт с максимальным номером как "исходящий".
        /// </summary>
        public static string FindBluetoothOutPort()
        {
            var ports = new List<BtPortInfo>();

            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT Name, PNPDeviceID FROM Win32_PnPEntity " +
                    "WHERE Name LIKE '%(COM%' AND PNPDeviceID LIKE 'BTHENUM%'"))
                {
                    foreach (ManagementObject mo in searcher.Get())
                    {
                        object nameObj = mo["Name"];
                        object pnpObj = mo["PNPDeviceID"];
                        if (nameObj == null || pnpObj == null)
                            continue;

                        string name = nameObj.ToString();
                        string pnp = pnpObj.ToString();

                        Match m = ComRx.Match(name);
                        if (!m.Success) continue;

                        int num;
                        if (!int.TryParse(m.Groups[1].Value.Substring(3), out num)) // "COM9" → "9"
                            continue;

                        var info = new BtPortInfo();
                        info.Name = name;
                        info.PnpDeviceId = pnp;
                        info.ComNumber = num;

                        ports.Add(info);
                    }
                }

                if (ports.Count == 0)
                    return null;

                // берём BT-порт с максимальным номером (обычно OUTGOING)
                BtPortInfo best = null;
                int bestNum = -1;

                foreach (BtPortInfo p in ports)
                {
                    if (p.ComNumber > bestNum)
                    {
                        best = p;
                        bestNum = p.ComNumber;
                    }
                }

                if (best == null)
                    return null;

                return "COM" + best.ComNumber.ToString();
            }
            catch
            {
                return null;
            }
        }
    }
}
