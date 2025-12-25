using System;
using System.Text;
using System.Drawing;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
//*****************************************************************************************
//                           LICENSE INFORMATION
//*****************************************************************************************
//   PCCom.SerialCommunication Version 1.0.0.0
//   Class file for managing serial port communication
//
//   Copyright (C) 2007  
//   Richard L. McCutchen 
//   Email: richard@psychocoder.net
//   Created: 20OCT07
//
//   This program is free software: you can redistribute it and/or modify
//   it under the terms of the GNU General Public License as published by
//   the Free Software Foundation, either version 3 of the License, or
//   (at your option) any later version.
//
//   This program is distributed in the hope that it will be useful,
//   but WITHOUT ANY WARRANTY; without even the implied warranty of
//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//   GNU General Public License for more details.
//
//   You should have received a copy of the GNU General Public License
//   along with this program.  If not, see <http://www.gnu.org/licenses/>.
//*****************************************************************************************
namespace DFlasher
{
    class CommunicationManager
    {
        #region Manager Enums
        /// <summary>
        /// enumeration to hold our transmission types
        /// </summary>
        public enum TransmissionType { Text, Hex }

        /// <summary>
        /// enumeration to hold our message types
        /// </summary>
        public enum MessageType { Incoming, Outgoing, Normal, Warning, Error };
        #endregion

        #region Manager Variables
        //property variables
        private string _baudRate = string.Empty;
        private string _parity = string.Empty;
        private string _stopBits = string.Empty;
        private string _dataBits = string.Empty;
        private string _portName = string.Empty;
        private TransmissionType _transType;
        private RichTextBox _displayWindow;
        private TextBox _TextBoxCtrl1;
        private RadioButton _rdBtnCtrl1;
        private RadioButton _rdBtnCtrl2;
        public bool ComPortIsOpen = false;
        //global manager variables
        private Color[] MessageColor = { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red };
        private SerialPort comPort = new SerialPort();
        private byte[] ReceivBuf = new byte[1024];
        private int CntBytes;
        private const int PACKET_LENGTH = 6;
        private byte[] ArMask = { 0xD, 0xD, 0, 0, 0xA, 0xA };
        private byte[] ArCmd = { 0xD, 0xE, 0, 0, 0xA, 0xA };//ěŕńęŕ îňâĺňíîăî ďŕęĺňŕ ńî çíŕ÷ĺíčĺě áŕéňŕ íŕďđŕâëĺíč˙ âđŕůĺíč˙

        // НОВОЕ: таймаут для отправки
        private const int WRITE_TIMEOUT_MS = 300;
        #endregion

        #region Manager Properties
        /// <summary>
        /// Property to hold the BaudRate
        /// of our manager class
        /// </summary>
        public string BaudRate
        {
            get { return _baudRate; }
            set { _baudRate = value; }
        }

        /// <summary>
        /// property to hold the Parity
        /// of our manager class
        /// </summary>
        public string Parity
        {
            get { return _parity; }
            set { _parity = value; }
        }

        /// <summary>
        /// property to hold the StopBits
        /// of our manager class
        /// </summary>
        public string StopBits
        {
            get { return _stopBits; }
            set { _stopBits = value; }
        }

        /// <summary>
        /// property to hold the DataBits
        /// of our manager class
        /// </summary>
        public string DataBits
        {
            get { return _dataBits; }
            set { _dataBits = value; }
        }

        /// <summary>
        /// property to hold the PortName
        /// of our manager class
        /// </summary>
        public string PortName
        {
            get { return _portName; }
            set { _portName = value; }
        }

        /// <summary>
        /// property to hold our TransmissionType
        /// of our manager class
        /// </summary>
        public TransmissionType CurrentTransmissionType
        {
            get { return _transType; }
            set { _transType = value; }
        }

        /// <summary>
        /// property to hold our display window
        /// value
        /// </summary>
        public RichTextBox DisplayWindow
        {
            get { return _displayWindow; }
            set { _displayWindow = value; }
        }
        public TextBox DisplayCtrl1
        {
            get { return _TextBoxCtrl1; }
            set { _TextBoxCtrl1 = value; }
        }
        public RadioButton DisplayCtrl2
        {
            get { return _rdBtnCtrl1; }
            set { _rdBtnCtrl1 = value; }
        }
        public RadioButton DisplayCtrl3
        {
            get { return _rdBtnCtrl2; }
            set { _rdBtnCtrl2 = value; }
        }
        #endregion

        #region Manager Constructors
        /// <summary>
        /// Constructor to set the properties of our Manager Class
        /// </summary>
        /// <param name="baud">Desired BaudRate</param>
        /// <param name="par">Desired Parity</param>
        /// <param name="sBits">Desired StopBits</param>
        /// <param name="dBits">Desired DataBits</param>
        /// <param name="name">Desired PortName</param>
        public CommunicationManager(string baud, string par, string sBits, string dBits, string name, RichTextBox rtb)
        {
            _baudRate = baud;
            _parity = par;
            _stopBits = sBits;
            _dataBits = dBits;
            _portName = name;
            _displayWindow = rtb;
            //now add an event handler
            comPort.DataReceived += new SerialDataReceivedEventHandler(comPort_DataReceived);
        }

        /// <summary>
        /// Comstructor to set the properties of our
        /// serial port communicator to nothing
        /// </summary>
        public CommunicationManager()
        {
            _baudRate = string.Empty;
            _parity = string.Empty;
            _stopBits = string.Empty;
            _dataBits = string.Empty;
            _portName = "COM1";
            _displayWindow = null;
            //add event handler
            comPort.DataReceived += new SerialDataReceivedEventHandler(comPort_DataReceived);
        }
        #endregion

        #region WriteData
        /// <summary>
        /// Безопасная отправка данных с защитой от зависания
        /// Возвращает true если отправка удалась, false если нет
        /// </summary>
        public bool WriteData(string msg)
        {
            try
            {
                // Устанавливаем таймаут перед отправкой
                comPort.WriteTimeout = WRITE_TIMEOUT_MS;

                if (!comPort.IsOpen)
                {
                    ComPortIsOpen = false;
                    SafeDisplayData(MessageType.Error, $"Порт не открыт: {msg}\n");
                    return false;
                }

                switch (CurrentTransmissionType)
                {
                    case TransmissionType.Text:
                        // ensure there's a delimiter so Arduino processes the command
                        if (!msg.EndsWith("\n") && !msg.EndsWith("\r") && !msg.EndsWith(";") && !msg.EndsWith(" "))
                        {
                            msg += "\n";
                        }

                        comPort.Write(msg);
                        SafeDisplayData(MessageType.Outgoing, msg);
                        return true;

                    case TransmissionType.Hex:
                        try
                        {
                            byte[] newMsg = HexToByte(msg);
                            comPort.Write(newMsg, 0, newMsg.Length);
                            SafeDisplayData(MessageType.Outgoing, ByteToHex(newMsg) + "\n");
                            return true;
                        }
                        catch (FormatException ex)
                        {
                            SafeDisplayData(MessageType.Error, ex.Message);
                            return false;
                        }

                    default:
                        if (!msg.EndsWith("\n") && !msg.EndsWith("\r") && !msg.EndsWith(";") && !msg.EndsWith(" "))
                        {
                            msg += "\n";
                        }

                        comPort.Write(msg);
                        SafeDisplayData(MessageType.Outgoing, msg);
                        return true;
                }
            }
            catch (TimeoutException)
            {
                // ТАЙМАУТ - порт завис
                ComPortIsOpen = false;
                SafeDisplayData(MessageType.Error, $"Таймаут отправки: {msg}\n");
                return false;
            }
            catch (InvalidOperationException)
            {
                // Порт закрылся во время отправки
                ComPortIsOpen = false;
                SafeDisplayData(MessageType.Error, $"Порт закрыт: {msg}\n");
                return false;
            }
            catch (System.IO.IOException)
            {
                // Ошибка ввода-вывода (устройство отключили)
                ComPortIsOpen = false;
                SafeDisplayData(MessageType.Error, $"Ошибка ввода-вывода: {msg}\n");
                return false;
            }
            catch (Exception ex)
            {
                SafeDisplayData(MessageType.Error, $"Ошибка отправки: {ex.Message}\n");
                ComPortIsOpen = false;
                return false;
            }
            finally
            {
                // Возвращаем бесконечный таймаут для операций чтения
                comPort.WriteTimeout = SerialPort.InfiniteTimeout;
            }
        }
        #endregion

        #region HexToByte
        /// <summary>
        /// method to convert hex string into a byte array
        /// </summary>
        /// <param name="msg">string to convert</param>
        /// <returns>a byte array</returns>
        private byte[] HexToByte(string msg)
        {
            //remove any spaces from the string
            msg = msg.Replace(" ", "");
            //create a byte array the length of the
            //divided by 2 (Hex is 2 characters in length)
            byte[] comBuffer = new byte[msg.Length / 2];
            //loop through the length of the provided string
            for (int i = 0; i < msg.Length; i += 2)
                //convert each set of 2 characters to a byte
                //and add to the array
                comBuffer[i / 2] = (byte)Convert.ToByte(msg.Substring(i, 2), 16);
            //return the array
            return comBuffer;
        }
        #endregion

        #region ByteToHex
        /// <summary>
        /// method to convert a byte array into a hex string
        /// </summary>
        /// <param name="comByte">byte array to convert</param>
        /// <returns>a hex string</returns>
        private string ByteToHex(byte[] comByte)
        {
            //create a new StringBuilder object
            StringBuilder builder = new StringBuilder(comByte.Length * 3);
            //loop through each byte in the array
            foreach (byte data in comByte)
                //convert the byte to a string and add to the stringbuilder
                builder.Append(Convert.ToString(data, 16).PadLeft(2, '0').PadRight(3, ' '));
            //return the converted value
            return builder.ToString().ToUpper();
        }
        #endregion

        #region DisplayData (безопасная версия)
        /// <summary>
        /// method to display the data to & from the port
        /// on the screen
        /// </summary>
        [STAThread]
        private void DisplayData(MessageType type, string msg)
        {
            _displayWindow.Invoke(new EventHandler(delegate
            {
                _displayWindow.SelectedText = string.Empty;
                _displayWindow.SelectionFont = new Font(_displayWindow.SelectionFont, FontStyle.Bold);
                _displayWindow.SelectionColor = MessageColor[(int)type];
                _displayWindow.AppendText(msg);
                _displayWindow.ScrollToCaret();
            }));
        }

        /// <summary>
        /// Безопасный вывод в UI - не блокирует при ошибках
        /// </summary>
        private void SafeDisplayData(MessageType type, string msg)
        {
            if (_displayWindow == null || _displayWindow.IsDisposed)
                return;

            if (_displayWindow.InvokeRequired)
            {
                try
                {
                    // Используем BeginInvoke вместо Invoke - не блокирует
                    _displayWindow.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (!_displayWindow.IsDisposed)
                            {
                                _displayWindow.SelectedText = string.Empty;
                                _displayWindow.SelectionFont = new Font(_displayWindow.SelectionFont, FontStyle.Bold);
                                _displayWindow.SelectionColor = MessageColor[(int)type];
                                _displayWindow.AppendText(msg);
                                _displayWindow.ScrollToCaret();
                            }
                        }
                        catch { /* Игнорируем ошибки UI */ }
                    }));
                }
                catch { /* Игнорируем ошибки BeginInvoke */ }
            }
            else
            {
                try
                {
                    _displayWindow.SelectedText = string.Empty;
                    _displayWindow.SelectionFont = new Font(_displayWindow.SelectionFont, FontStyle.Bold);
                    _displayWindow.SelectionColor = MessageColor[(int)type];
                    _displayWindow.AppendText(msg);
                    _displayWindow.ScrollToCaret();
                }
                catch { /* Игнорируем ошибки */ }
            }
        }
        #endregion

        #region DisplayData2 (безопасная версия)
        [STAThread]
        private void DisplayData2(MessageType type, string msg)
        {
            _TextBoxCtrl1.Invoke(new EventHandler(delegate
            {
                _TextBoxCtrl1.ForeColor = MessageColor[(int)type];
                msg = msg.Replace(" ", "");
                if (msg != "")
                {
                    int decValue = Convert.ToInt32(msg, 16);
                    _TextBoxCtrl1.Text = decValue.ToString();
                }
            }));
        }

        private void SafeDisplayData2(MessageType type, string msg)
        {
            if (_TextBoxCtrl1 == null || _TextBoxCtrl1.IsDisposed)
                return;

            if (_TextBoxCtrl1.InvokeRequired)
            {
                try
                {
                    _TextBoxCtrl1.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (!_TextBoxCtrl1.IsDisposed)
                            {
                                _TextBoxCtrl1.ForeColor = MessageColor[(int)type];
                                msg = msg.Replace(" ", "");
                                if (msg != "")
                                {
                                    int decValue = Convert.ToInt32(msg, 16);
                                    _TextBoxCtrl1.Text = decValue.ToString();
                                }
                            }
                        }
                        catch { }
                    }));
                }
                catch { }
            }
            else
            {
                try
                {
                    _TextBoxCtrl1.ForeColor = MessageColor[(int)type];
                    msg = msg.Replace(" ", "");
                    if (msg != "")
                    {
                        int decValue = Convert.ToInt32(msg, 16);
                        _TextBoxCtrl1.Text = decValue.ToString();
                    }
                }
                catch { }
            }
        }
        #endregion

        #region DisplayData3 (безопасная версия)
        [STAThread]
        private void DisplayData3(MessageType type, string msg)
        {
            _rdBtnCtrl1.Invoke(new EventHandler(delegate
            {
                msg = msg.Replace(" ", "");
                if (msg != "")
                {
                    int decValue = Convert.ToInt32(msg, 16);
                    GVars.LockCtrl = true;
                    _rdBtnCtrl1.Checked = decValue == 1; //rbtDirSouth
                    GVars.LockCtrl = false;
                }
            }));
        }

        private void SafeDisplayData3(MessageType type, string msg)
        {
            if (_rdBtnCtrl1 == null || _rdBtnCtrl1.IsDisposed)
                return;

            if (_rdBtnCtrl1.InvokeRequired)
            {
                try
                {
                    _rdBtnCtrl1.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (!_rdBtnCtrl1.IsDisposed)
                            {
                                msg = msg.Replace(" ", "");
                                if (msg != "")
                                {
                                    int decValue = Convert.ToInt32(msg, 16);
                                    GVars.LockCtrl = true;
                                    _rdBtnCtrl1.Checked = decValue == 1;
                                    GVars.LockCtrl = false;
                                }
                            }
                        }
                        catch { }
                    }));
                }
                catch { }
            }
            else
            {
                try
                {
                    msg = msg.Replace(" ", "");
                    if (msg != "")
                    {
                        int decValue = Convert.ToInt32(msg, 16);
                        GVars.LockCtrl = true;
                        _rdBtnCtrl1.Checked = decValue == 1;
                        GVars.LockCtrl = false;
                    }
                }
                catch { }
            }
        }
        #endregion

        #region DisplayData4 (безопасная версия)
        [STAThread]
        private void DisplayData4(MessageType type, string msg)
        {
            _rdBtnCtrl2.Invoke(new EventHandler(delegate
            {
                msg = msg.Replace(" ", "");
                if (msg != "")
                {
                    int decValue = Convert.ToInt32(msg, 16);
                    GVars.LockCtrl = true;
                    _rdBtnCtrl2.Checked = decValue == 0; //rbtDirNorth
                    GVars.LockCtrl = false;
                }
            }));
        }

        private void SafeDisplayData4(MessageType type, string msg)
        {
            if (_rdBtnCtrl2 == null || _rdBtnCtrl2.IsDisposed)
                return;

            if (_rdBtnCtrl2.InvokeRequired)
            {
                try
                {
                    _rdBtnCtrl2.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            if (!_rdBtnCtrl2.IsDisposed)
                            {
                                msg = msg.Replace(" ", "");
                                if (msg != "")
                                {
                                    int decValue = Convert.ToInt32(msg, 16);
                                    GVars.LockCtrl = true;
                                    _rdBtnCtrl2.Checked = decValue == 0;
                                    GVars.LockCtrl = false;
                                }
                            }
                        }
                        catch { }
                    }));
                }
                catch { }
            }
            else
            {
                try
                {
                    msg = msg.Replace(" ", "");
                    if (msg != "")
                    {
                        int decValue = Convert.ToInt32(msg, 16);
                        GVars.LockCtrl = true;
                        _rdBtnCtrl2.Checked = decValue == 0;
                        GVars.LockCtrl = false;
                    }
                }
                catch { }
            }
        }
        #endregion

        #region OpenPort
        public bool OpenPort()
        {
            try
            {
                //first check if the port is already open
                //if its open then close it
                if (comPort.IsOpen == true) comPort.Close();

                //set the properties of our SerialPort Object
                comPort.BaudRate = int.Parse(_baudRate);    //BaudRate
                comPort.DataBits = int.Parse(_dataBits);    //DataBits
                comPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), _stopBits);    //StopBits
                comPort.Parity = (Parity)Enum.Parse(typeof(Parity), _parity);    //Parity
                comPort.PortName = _portName;   //PortName
                comPort.DtrEnable = true;

                // Устанавливаем таймауты
                comPort.ReadTimeout = 1000;
                comPort.WriteTimeout = WRITE_TIMEOUT_MS;

                //now open the port
                comPort.Open();
                ComPortIsOpen = true;

                //display message
                SafeDisplayData(MessageType.Normal, "Port opened at " + DateTime.Now + "\n");

                //return true
                return true;
            }
            catch (Exception ex)
            {
                SafeDisplayData(MessageType.Error, ex.Message);
                ComPortIsOpen = false;
                return false;
            }
        }
        #endregion

        public bool ClosePort()
        {
            try
            {
                if (comPort.IsOpen == true)
                {
                    // Сбрасываем управляющие линии перед закрытием
                    comPort.DtrEnable = false;
                    comPort.RtsEnable = false;
                    Thread.Sleep(50);
                    comPort.Close();
                }
                ComPortIsOpen = false;
                SafeDisplayData(MessageType.Normal, "Port closed\n");
                return true;
            }
            catch (Exception ex)
            {
                SafeDisplayData(MessageType.Error, ex.Message);
                ComPortIsOpen = false;
                return false;
            }
        }

        /// <summary>
        /// Принудительное закрытие порта при зависании
        /// </summary>
        public void ForceClosePort()
        {
            try
            {
                ComPortIsOpen = false;
                if (comPort.IsOpen)
                {
                    // Быстро сбрасываем линии
                    comPort.DtrEnable = false;
                    comPort.RtsEnable = false;

                    // Пытаемся закрыть в отдельном потоке
                    Thread closeThread = new Thread(() =>
                    {
                        try { comPort.Close(); } catch { }
                    });
                    closeThread.Start();

                    // Даем 100мс на закрытие
                    if (!closeThread.Join(100))
                    {
                        closeThread.Interrupt();
                    }
                }
            }
            catch { /* Игнорируем все ошибки */ }
        }

        #region SetParityValues
        public void SetParityValues(object obj)
        {
            foreach (string str in Enum.GetNames(typeof(Parity)))
            {
                ((ComboBox)obj).Items.Add(str);
            }
        }
        #endregion

        #region SetStopBitValues
        public void SetStopBitValues(object obj)
        {
            foreach (string str in Enum.GetNames(typeof(StopBits)))
            {
                ((ComboBox)obj).Items.Add(str);
            }
        }
        #endregion

        #region SetPortNameValues
        public void SetPortNameValues(object obj)
        {

            foreach (string str in SerialPort.GetPortNames())
            {
                ((ComboBox)obj).Items.Add(str);
            }
        }
        #endregion

        bool ChckSyncMask(byte[] Rbuf, byte[] ArMsk)
        {
            return (Rbuf[0] == ArMsk[0]) /*&& Rbuf[1] == ArMsk[1]*/ && (Rbuf[4] == ArMsk[4]) && (Rbuf[5] == ArMsk[5]);
        }
        bool ChckCmd(byte[] Rbuf, byte[] ArMsk)
        {
            return (Rbuf[1] == ArMsk[1]);
        }

        #region comPort_DataReceived
        /// <summary>
        /// method that will be called when theres data waiting in the buffer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void comPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //determine the mode the user selected (binary/string)
            switch (CurrentTransmissionType)
            {
                //user chose string
                case TransmissionType.Text:
                    //read data waiting in the buffer
                    string msg = comPort.ReadExisting();
                    //display the data to the user
                    SafeDisplayData(MessageType.Incoming, msg + "\n");
                    break;
                //user chose binary
                case TransmissionType.Hex:
                    //retrieve number of bytes in the buffer
                    int bytes = comPort.BytesToRead;
                    //create a byte array to hold the awaiting data
                    byte[] comBuffer = new byte[bytes];
                    //read the data and store it
                    comPort.Read(comBuffer, 0, bytes);

                    Array.Copy(comBuffer, 0, ReceivBuf, CntBytes, bytes);//ń ęŕćäűě âűçîâîě ďĺđĺíîńčě äŕííűĺ â áóôĺđ íŕęîďëĺíč˙
                    CntBytes += bytes;
                    while (CntBytes >= PACKET_LENGTH)
                    {
                        if (ChckSyncMask(ReceivBuf, ArMask))//ďűňŕĺěń˙ ńčíőđîíčçčđîâŕňüń˙ - íŕéňč ďđčçíŕęč ďŕęĺňŕ
                        {
                            if (ChckCmd(ReceivBuf, ArCmd))//ďđčřëî çíŕ÷ĺíčĺ íŕďđŕâëĺíč˙ âđŕůĺíč˙
                            {
                                byte[] buf2 = new byte[1];
                                Array.Copy(ReceivBuf, 2, buf2, 0, 1);//3-é áŕéň DirRotationClockWise = 1, 4-é âńĺăäŕ 0
                                SafeDisplayData3(MessageType.Incoming, ByteToHex(buf2));//îňîáđŕćŕĺě íŕ ęîíňđîëŕő _rdBtnCtrl1 - rbtDirNorth
                                SafeDisplayData4(MessageType.Incoming, ByteToHex(buf2));//_rdBtnCtrl2 - rbtDirSouth
                            }
                            else
                            {
                                byte[] buf2 = new byte[2];
                                Array.Copy(ReceivBuf, 2, buf2, 0, 2);
                                SafeDisplayData2(MessageType.Incoming, ByteToHex(buf2));
                            }
                            Array.Copy(ReceivBuf, PACKET_LENGTH, ReceivBuf, 0, CntBytes);//óäŕë˙ĺě čç íŕęîďëĺííîăî áóôĺđŕ áŕéňű îáđŕáîňŕííîăî âűřĺ ďŕęĺňŕ
                            CntBytes -= PACKET_LENGTH;
                        }
                        else// ĺńëč íĺ óäŕëîńü, ńäâčăŕĺě áóôĺđ íŕ 1 áŕéň
                        {
                            CntBytes--;
                            Array.Copy(ReceivBuf, 1, ReceivBuf, 0, bytes);
                        }
                    }

                    //display the data to the user
                    SafeDisplayData(MessageType.Incoming, ByteToHex(comBuffer) + "\n");
                    break;
                default:
                    //read data waiting in the buffer
                    string str = comPort.ReadExisting();
                    //display the data to the user
                    SafeDisplayData(MessageType.Incoming, str + "\n");
                    break;
            }
        }
        #endregion
    }
}