using System;
using System.Text;
using System.Drawing;
using System.IO.Ports;
using System.Windows.Forms;
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
        private byte[] ArMask = {0xD, 0xD, 0, 0, 0xA, 0xA};
        private byte[] ArCmd = { 0xD, 0xE, 0, 0, 0xA, 0xA };//маска ответного пакета со значением байта направления вращения
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
        public void WriteData(string msg)
        {
            try
            {
                if (!comPort.IsOpen)
                    comPort.Open();

                switch (CurrentTransmissionType)
                {
                    case TransmissionType.Text:
                        // ensure there's a delimiter so Arduino processes the command
                        if (!msg.EndsWith("\n") && !msg.EndsWith("\r") && !msg.EndsWith(";") && !msg.EndsWith(" "))
                        {
                            msg += "\n";
                        }

                        comPort.Write(msg);
                        DisplayData(MessageType.Outgoing, msg); // already has newline if we added
                        break;

                    case TransmissionType.Hex:
                        try
                        {
                            byte[] newMsg = HexToByte(msg);
                            comPort.Write(newMsg, 0, newMsg.Length);
                            DisplayData(MessageType.Outgoing, ByteToHex(newMsg) + "\n");
                        }
                        catch (FormatException ex)
                        {
                            DisplayData(MessageType.Error, ex.Message);
                        }
                        break;

                    default:
                        if (!msg.EndsWith("\n") && !msg.EndsWith("\r") && !msg.EndsWith(";") && !msg.EndsWith(" "))
                        {
                            msg += "\n";
                        }

                        comPort.Write(msg);
                        DisplayData(MessageType.Outgoing, msg);
                        break;
                }
            }
            catch (Exception ex)
            {
                DisplayData(MessageType.Error, $"WriteData failed: {ex.Message}");
            }
            finally
            {
                // keep focus / selection behavior if needed
                //_displayWindow.SelectAll();
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

        #region DisplayData
        /// <summary>
        /// method to display the data to & from the port
        /// on the screen
        /// </summary>
        /// <param name="type">MessageType of the message</param>
        /// <param name="msg">Message to display</param>
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
        #endregion

        #region DisplayData2
        [STAThread]
        private void DisplayData2(MessageType type, string msg)
        {
            _TextBoxCtrl1.Invoke(new EventHandler(delegate
            {
                _TextBoxCtrl1.ForeColor = MessageColor[(int)type];
                msg = msg.Replace(" ", "");
                if (msg != "") {
                    int decValue = Convert.ToInt32(msg, 16);
                    _TextBoxCtrl1.Text = decValue.ToString();
                }
            }));
        }
        #endregion

        #region DisplayData3
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
        #endregion

        #region DisplayData4
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
                //now open the port
                comPort.Open();
                ComPortIsOpen = true;
                //display message
                DisplayData(MessageType.Normal, "Port opened at " + DateTime.Now + "\n");
                //return true
                return true;
            }
            catch (Exception ex)
            {
                DisplayData(MessageType.Error, ex.Message);
                return false;
            }
        }
        #endregion

        public bool ClosePort()
        {
            try
            {
                if (comPort.IsOpen == true) comPort.Close();
                ComPortIsOpen = false;
                return true;
            }
            catch (Exception ex)
            {
                DisplayData(MessageType.Error, ex.Message);
                return false;
            }
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
                    DisplayData(MessageType.Incoming, msg + "\n");
                    break;
                //user chose binary
                case TransmissionType.Hex:
                    //retrieve number of bytes in the buffer
                    int bytes = comPort.BytesToRead;
                    //create a byte array to hold the awaiting data
                    byte[] comBuffer = new byte[bytes];
                    //read the data and store it
                    comPort.Read(comBuffer, 0, bytes);

                    Array.Copy(comBuffer, 0, ReceivBuf, CntBytes, bytes);//с каждым вызовом переносим данные в буфер накопления
                    CntBytes += bytes;
                    while (CntBytes >= PACKET_LENGTH)
                    {
                        if (ChckSyncMask(ReceivBuf, ArMask))//пытаемся синхронизироваться - найти признаки пакета
                        {
                            if (ChckCmd(ReceivBuf, ArCmd))//пришло значение направления вращения
                            {
                                byte[] buf2 = new byte[1];
                                Array.Copy(ReceivBuf, 2, buf2, 0, 1);//3-й байт DirRotationClockWise = 1, 4-й всегда 0
                                DisplayData3(MessageType.Incoming, ByteToHex(buf2));//отображаем на контролах _rdBtnCtrl1 - rbtDirNorth
                                DisplayData4(MessageType.Incoming, ByteToHex(buf2));//_rdBtnCtrl2 - rbtDirSouth
                            }
                            else
                            {
                                byte[] buf2 = new byte[2];
                                Array.Copy(ReceivBuf, 2, buf2, 0, 2);
                                DisplayData2(MessageType.Incoming, ByteToHex(buf2));
                            }
                            Array.Copy(ReceivBuf, PACKET_LENGTH, ReceivBuf, 0, CntBytes);//удаляем из накопленного буфера байты обработанного выше пакета
                            CntBytes -= PACKET_LENGTH;
                        }
                        else// если не удалось, сдвигаем буфер на 1 байт
                        {
                            CntBytes--;
                            Array.Copy(ReceivBuf, 1, ReceivBuf, 0, bytes);
                        }
                    }

                    //display the data to the user
                    DisplayData(MessageType.Incoming, ByteToHex(comBuffer) + "\n");
                    break;
                default:
                    //read data waiting in the buffer
                    string str = comPort.ReadExisting();
                    //display the data to the user
                    DisplayData(MessageType.Incoming, str + "\n");
                    break;
            }
        }
        #endregion
    }
}
