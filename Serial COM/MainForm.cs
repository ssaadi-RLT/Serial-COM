using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO.Ports;
using System.Drawing;
using System.Management;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Megamind.IO.Serial;

namespace Serial_COM
{
    public partial class MainForm : Form
    {
        #region enum ans const

        public const string SERIAL_ENTITY_CLASS_GUID = "{4d36e978-e325-11ce-bfc1-08002be10318}";
        public static readonly string[] StddBaudRate = { "9600", "14400", "19200", "38400", "57600", "115200" };

        public enum EncodeType { Auto, Ascii, Hex }

        public class WmiSerialPortInfo
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string Manufacturer { get; set; }
            public string Details
            {
                get => string.Format("{0} - {1} | {2}", Name, Description, Manufacturer);
            }
        }

        #endregion

        #region Data

        bool _isConnected = false;
        bool _sendOnKeyPress = false;
        bool _appendCrOnTx = false;
        bool _appendNlOnTx = false;
        bool _appendNlOnRx = false;
        bool _appendTsOnTx = true;
        bool _appendTsOnRx = false;
        bool _setDtrOnConnect = false;
        bool _showWmiPortNames = false;

        SerialCom _serialCom = new SerialCom("COM1", 9600);
        EncodeType _txEncode = EncodeType.Ascii;
        EncodeType _rxDecode = EncodeType.Ascii;

        int _lastSendIndex = 0;
        readonly List<string> _sendHistory = new List<string>();

        string[] _portNames = new string[0];
        string[] _portViewNames = new string[0];

        static ManagementEventWatcher _deviceArrivalWatcher;
        static ManagementEventWatcher _deviceRemovalWatcher;

        #endregion

        #region Internal Methods

        //thread safe log handler
        private void AppendEventLog(string str, Color? color = null, bool appendNewLine = true)
        {
            var clr = color ?? Color.Black;
            if (appendNewLine) str += Environment.NewLine;

            // update UI from dispatcher thread
            Invoke(new MethodInvoker(() =>
            {
                richTextBoxExEventLog.AppendText(str, clr);
            }));
        }

        private void MonitorDeviceChanges()
        {
            var deviceArrivalQuery = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2");
            var deviceRemovalQuery = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 3");
            _deviceArrivalWatcher = new ManagementEventWatcher(deviceArrivalQuery);
            _deviceRemovalWatcher = new ManagementEventWatcher(deviceRemovalQuery);
            _deviceArrivalWatcher.EventArrived += (o, args) => DeviceChangedCallback();
            _deviceRemovalWatcher.EventArrived += (sender, eventArgs) => DeviceChangedCallback();
            _deviceArrivalWatcher.Start();
            _deviceRemovalWatcher.Start();
        }

        //called several timse when device arrived or removed
        private void DeviceChangedCallback()
        {
            try
            {
                if (!_isConnected) UpdateSerialPortNames();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        public static IEnumerable<WmiSerialPortInfo> GetWmiSerialPortInfo()
        {
            var portNames = new List<WmiSerialPortInfo>();
            var wmiQuery = "SELECT * FROM Win32_PnPEntity WHERE ClassGuid='" + SERIAL_ENTITY_CLASS_GUID + "'";
            using (var searcher = new ManagementObjectSearcher("root\\CIMV2", wmiQuery))
            {
                foreach (var item in searcher.Get())
                {
                    var port = new WmiSerialPortInfo();
                    try
                    {
                        var allMatches = Regex.Matches(item["Name"].ToString(), @"\(COM(\d)+\)");
                        if (allMatches.Count > 0) port.Name = allMatches[0].Value.Substring(1, allMatches[0].Value.Length - 2);
                        port.Description = item["Description"].ToString();
                        port.Manufacturer = item["Manufacturer"].ToString();
                    }
                    catch { }
                    if (!String.IsNullOrEmpty(port.Name)) portNames.Add(port);
                }
            }
            return portNames.OrderBy(p => p.Name);
        }

        //update combobox if port name sequence changed
        private void UpdateSerialPortNames()
        {
            var portNames = new string[0];
            var portViewNames = new string[0];

            if (_showWmiPortNames)
            {
                var wmiPortNames = GetWmiSerialPortInfo();
                portNames = wmiPortNames.Select(p => p.Name).ToArray();
                portViewNames = wmiPortNames.Select(p => p.Details).ToArray();
            }
            else
            {
                portNames = SerialPort.GetPortNames();
                Array.Sort(portNames);
                portViewNames = portNames;
            }

            if (!_portNames.SequenceEqual(portNames) || !_portViewNames.SequenceEqual(portViewNames))
            {
                _portNames = portNames;
                _portViewNames = portViewNames;
                Invoke(new Action(() =>
                {
                    var prevSelection = comboBoxPortName.Text;
                    comboBoxPortName.ComboBox.DataSource = portViewNames;

                    if (comboBoxPortName.Items.Contains(prevSelection))
                        comboBoxPortName.SelectedIndex = comboBoxPortName.Items.IndexOf(prevSelection);
                    else if (comboBoxPortName.Items.Count > 0)
                        comboBoxPortName.SelectedIndex = comboBoxPortName.Items.Count - 1;
                }));
            }
        }

        public void UpdateSerialPortDescription()
        {
            var portName = comboBoxPortName.Text;
            if (comboBoxPortName.SelectedIndex >= 0)
                portName = _portNames[comboBoxPortName.SelectedIndex];
            var portInfo = GetWmiSerialPortInfo().FirstOrDefault(p => p.Name == portName);
            if (portInfo != null) toolStripPortDetails.Text = portInfo.Details;
            else toolStripPortDetails.Text = portName;
        }

        public static string ByteArrayToFormatedString(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (var item in bytes)
            {
                if (item == 10) sb.Append("<LF>");
                else if (item == 13) sb.Append("<CR>");
                else if (item < 32 || item > 126) sb.AppendFormat("<{0:X2}>", item);
                else sb.AppendFormat("{0}", (char)item);
            }
            return sb.ToString();
        }

        public static string ByteArrayToHexString(byte[] bytes, string separator = "")
        {
            return BitConverter.ToString(bytes).Replace("-", separator);
        }

        public static byte[] HexStringToByteArray(string hexstr)
        {
            hexstr.Trim();
            hexstr = hexstr.Replace("-", "");
            hexstr = hexstr.Replace(" ", "");
            return Enumerable.Range(0, hexstr.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hexstr.Substring(x, 2), 16))
                             .ToArray();
        }

        private void PopupException(string message, string caption = "Exception")
        {
            Invoke(new Action(() =>
            {
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }));
        }

        #endregion

        #region ctor

        public MainForm()
        {
            InitializeComponent();
            richTextBoxTxData.ForeColor = Color.DarkGreen;
            richTextBoxTxData.Font = new Font("Consolas", 9);
            richTextBoxExEventLog.Font = new Font("Consolas", 9);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            try
            {
                UpdateSerialPortNames();
                comboBoxBaudRate.ComboBox.DataSource = StddBaudRate;

                Config.Load();
                if (comboBoxPortName.Items.Contains(Config.PortName))
                    comboBoxPortName.SelectedIndex = comboBoxPortName.Items.IndexOf(Config.PortName);
                if (comboBoxBaudRate.Items.Contains(Config.BaudRate))
                    comboBoxBaudRate.SelectedIndex = comboBoxBaudRate.Items.IndexOf(Config.BaudRate);

                MonitorDeviceChanges();
                toolStripStatusLabel1.Text = "Ready to Connect";
                labelRxSelection.Text = "RxSel: 0";
                labelTxSelection.Text = "TxSel: 0";
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                var portName = comboBoxPortName.Text;
                if (comboBoxPortName.SelectedIndex >= 0)
                    portName = _portNames[comboBoxPortName.SelectedIndex];
                Config.PortName = portName;
                Config.BaudRate = comboBoxBaudRate.Text;
                Config.Save();
                _serialCom.Dispose();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region File Menu

        private void ToolStripMenuItemNew_Click(object sender, EventArgs e)
        {
            Process.Start(Application.ExecutablePath);
        }

        private void ToolStripMenuItemSaveAs_Click(object sender, EventArgs e)
        {
            try
            {
                using (var sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                    var name = DateTime.Today.ToLongDateString() + "_" + DateTime.Today.ToLongTimeString();
                    sfd.FileName = name.Replace(':', '-').Replace(',', '-');
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        richTextBoxExEventLog.SaveFile(sfd.FileName, RichTextBoxStreamType.PlainText);
                    }
                }
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ToolStripMenuItemAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Developer - GSM Rana\nwww.gsmrana.com", "About");
        }

        #endregion

        #region Send Menu

        private void ToolStripMenuItemTxEncodeAuto_Click(object sender, EventArgs e)
        {
            _txEncode = EncodeType.Auto;
            encodeASCIIToolStripMenuItemTx.Checked = false;
            encodeHEXToolStripMenuItemTx.Checked = false;
            encodeAutoToolStripMenuItemTx.Checked = true;
        }

        private void ToolStripMenuItemTxEncodeAscii_Click(object sender, EventArgs e)
        {
            _txEncode = EncodeType.Ascii;
            encodeAutoToolStripMenuItemTx.Checked = false;
            encodeHEXToolStripMenuItemTx.Checked = false;
            encodeASCIIToolStripMenuItemTx.Checked = true;
        }

        private void ToolStripMenuItemTxEncodeHEX_Click(object sender, EventArgs e)
        {
            _txEncode = EncodeType.Hex;
            encodeAutoToolStripMenuItemTx.Checked = false;
            encodeASCIIToolStripMenuItemTx.Checked = false;
            encodeHEXToolStripMenuItemTx.Checked = true;
        }

        private void ToolStripMenuItemTxAppendCR_CheckedChanged(object sender, EventArgs e)
        {
            _appendCrOnTx = toolStripMenuItemTxAppendCR.Checked;
        }

        private void ToolStripMenuItemTxAppendNL_CheckedChanged(object sender, EventArgs e)
        {
            _appendNlOnTx = toolStripMenuItemTxAppendNL.Checked;
        }

        private void ToolStripMenuItemTxAppendTs_CheckedChanged(object sender, EventArgs e)
        {
            _appendTsOnTx = toolStripMenuItemTxAppendTs.Checked;
        }

        private void ToolStripMenuItemTxSendOnKeyPress_Click(object sender, EventArgs e)
        {
            _sendOnKeyPress = sendOnKeyPressToolStripMenuItemTx.Checked;
            if (_sendOnKeyPress)
            {
                richTextBoxTxData.Clear();
                richTextBoxTxData.Focus();
            }
        }

        #endregion

        #region Receive Menu

        private void ToolStripMenuItemRxDecodeAuto_Click(object sender, EventArgs e)
        {
            _rxDecode = EncodeType.Auto;
            decodeASCIIToolStripMenuItemRx.Checked = false;
            decodeHEXToolStripMenuItemRx.Checked = false;
            decodeAutoToolStripMenuItemRx.Checked = true;
        }

        private void ToolStripMenuItemRxDecodeAscii_Click(object sender, EventArgs e)
        {
            _rxDecode = EncodeType.Ascii;
            decodeAutoToolStripMenuItemRx.Checked = false;
            decodeHEXToolStripMenuItemRx.Checked = false;
            decodeASCIIToolStripMenuItemRx.Checked = true;
        }

        private void ToolStripMenuItemRxDecodeHex_Click(object sender, EventArgs e)
        {
            _rxDecode = EncodeType.Hex;
            decodeAutoToolStripMenuItemRx.Checked = false;
            decodeASCIIToolStripMenuItemRx.Checked = false;
            decodeHEXToolStripMenuItemRx.Checked = true;
        }
        private void ToolStripMenuItemWordWrap_CheckedChanged(object sender, EventArgs e)
        {
            richTextBoxExEventLog.WordWrap = toolStripMenuItemWordWrap.Checked;
        }

        private void ToolStripMenuItemAutoScroll_CheckedChanged(object sender, EventArgs e)
        {
            richTextBoxExEventLog.Autoscroll = toolStripMenuItemAutoScroll.Checked;
        }

        private void ToolStripMenuItemRxAppendNL_CheckedChanged(object sender, EventArgs e)
        {
            _appendNlOnRx = toolStripMenuItemRxAppendNL.Checked;
        }

        private void ToolStripMenuItemRxAppendTS_CheckedChanged(object sender, EventArgs e)
        {
            _appendTsOnRx = toolStripMenuItemRxAppendTS.Checked;
        }

        private void ToolStripMenuItemRxAppendNL_Click(object sender, EventArgs e)
        {
            _appendNlOnRx = toolStripMenuItemRxAppendNL.Checked;
        }

        #endregion

        #region Tools Menu

        private void ToolStripMenuItemAwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = toolStripMenuItemAwaysOnTop.Checked;
        }

        private void ToolStripMenuItemToolsResetOnConnect_CheckedChanged(object sender, EventArgs e)
        {
            _setDtrOnConnect = toolStripMenuItemSetDtrOnConnect.Checked;
        }

        private void ToolStripMenuItemWmiPortNames_CheckedChanged(object sender, EventArgs e)
        {
            _showWmiPortNames = toolStripMenuItemWmiPortNames.Checked;
            try
            {
                if (_showWmiPortNames) comboBoxPortName.Width = 350;
                else comboBoxPortName.Width = 100;
                UpdateSerialPortNames();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region ToolStrip Events

        private void ComboBoxPortList_DropDown(object sender, EventArgs e)
        {
            try
            {
                UpdateSerialPortNames();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ComboBoxPortName_DropDownClosed(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ComboBoxPortName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                UpdateSerialPortDescription();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isConnected)
                {
                    var portName = comboBoxPortName.Text;
                    if (comboBoxPortName.SelectedIndex >= 0)
                        portName = _portNames[comboBoxPortName.SelectedIndex];
                    var baudRate = int.Parse(comboBoxBaudRate.Text);

                    _serialCom = new SerialCom(portName, baudRate, _setDtrOnConnect);
                    _serialCom.OnDataReceived += SerialCom_OnDataReceived;
                    _serialCom.OnException += SerialCom_OnException;
                    _serialCom.Open();

                    _isConnected = true;
                    buttonConnect.Text = "Disconnect";
                    this.Text = string.Format("{0} [{1} - {2}]", Application.ProductName, portName, baudRate);
                    toolStripStatusLabel1.Text = "Connected at Baudrate " + baudRate;
                    comboBoxPortName.Enabled = false;
                    comboBoxBaudRate.Enabled = false;
                }
                else
                {
                    _isConnected = false;
                    buttonConnect.Text = "Connect";
                    comboBoxPortName.Enabled = true;
                    comboBoxBaudRate.Enabled = true;
                    toolStripStatusLabel1.Text = "Disconnected";
                    this.Text = Application.ProductName;
                    _serialCom.Close();
                }
                UpdateSerialPortDescription();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonList_Click(object sender, EventArgs e)
        {
            try
            {
                using (var tableview = new TableViwForm())
                {
                    tableview.Tittle = "Serial Port List";
                    tableview.ColumnHeaders.Add(new ColumnHeader("Name", 80));
                    tableview.ColumnHeaders.Add(new ColumnHeader("WMI Info", 200));
                    tableview.ColumnHeaders.Add(new ColumnHeader("Manufacturer", 180));
                    foreach (var item in GetWmiSerialPortInfo())
                        tableview.DataRows.Add(new[] { item.Name, item.Description, item.Manufacturer });
                    tableview.ShowDialog();
                    var idx = tableview.SelectedIndex;
                    if (idx >= 0 && comboBoxPortName.Items.Count > idx)
                        comboBoxPortName.SelectedIndex = idx;
                }
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonCopyText_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBoxExEventLog.Text.Length > 0)
                {
                    Clipboard.SetText(richTextBoxExEventLog.Text);
                }
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ToolStripButtonClear_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBoxExEventLog.Clear();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region RichTextBox Events

        private void RichTextBoxExEventLog_SelectionChanged(object sender, EventArgs e)
        {
            labelRxSelection.Text = string.Format("RxSel: {0}", richTextBoxExEventLog.SelectionLength);
        }

        private void RichTextBoxTxData_SelectionChanged(object sender, EventArgs e)
        {
            labelTxSelection.Text = string.Format("TxSel: {0}", richTextBoxTxData.SelectionLength);
        }

        #endregion

        #region Rx Context Menu

        private void ContextMenuStripRx_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            copyToolStripMenuItemRx.Enabled = richTextBoxExEventLog.SelectionLength > 0;
            copyAllToolStripMenuItemRx.Enabled = richTextBoxExEventLog.Text.Length > 0;
        }

        private void CopyToolStripMenuItemRx_Click(object sender, EventArgs e)
        {
            richTextBoxExEventLog.Copy();
        }

        private void CopyAllToolStripMenuItemRx_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBoxExEventLog.Text);
        }

        #endregion

        #region Tx Context Menu

        private void ContextMenuStripTx_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            cutToolStripMenuItemTx.Enabled = richTextBoxTxData.SelectionLength > 0;
            copyToolStripMenuItemTx.Enabled = richTextBoxTxData.SelectionLength > 0;
        }

        private void CutToolStripMenuItemTx_Click(object sender, EventArgs e)
        {
            richTextBoxTxData.Cut();
        }

        private void CopyToolStripMenuItemTx_Click(object sender, EventArgs e)
        {
            richTextBoxTxData.Copy();
        }

        private void PasteToolStripMenuItemTx_Click(object sender, EventArgs e)
        {
            richTextBoxTxData.Paste();
        }

        private void SelectAllToolStripMenuItemTx_Click(object sender, EventArgs e)
        {
            richTextBoxTxData.SelectAll();
        }

        #endregion

        #region Sending Data

        private void RichTextBoxTxData_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control && e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                }

                if (e.KeyCode == Keys.Up)
                {
                    if (_lastSendIndex > 0) _lastSendIndex--;
                    richTextBoxTxData.Text = _sendHistory[_lastSendIndex];
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Down)
                {
                    if (_lastSendIndex < _sendHistory.Count - 1) _lastSendIndex++;
                    richTextBoxTxData.Text = _sendHistory[_lastSendIndex];
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void RichTextBoxTxData_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void RichTextBoxTxData_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!_isConnected) return;

            try
            {
                if (_sendOnKeyPress)
                {
                    var keyval = (byte)e.KeyChar;
                    _serialCom.Write(new byte[] { keyval });

                    var txstr = "";
                    if (keyval > 32 && keyval < 127) txstr = e.KeyChar.ToString();
                    else txstr = "0x" + keyval.ToString("X2");
                    AppendEventLog(txstr, Color.DarkGreen);
                }
                else if (e.KeyChar == '\n') //CTRL+ENTER
                {
                    e.Handled = true;
                    buttonSend.PerformClick();
                }
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void ButtonSend_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isConnected) throw new Exception("Please Connect First!");

                var txstr = richTextBoxTxData.Text;
                if (string.IsNullOrEmpty(txstr)) return;

                //replace windows default newline char \n with \r
                txstr = txstr.Replace("\r\n", "\n");
                txstr = txstr.Replace('\n', '\r');

                _lastSendIndex = _sendHistory.Count;
                var idx = _sendHistory.IndexOf(txstr);
                if (idx >= 0) //already exist in history
                    _sendHistory.RemoveAt(idx);
                _sendHistory.Add(txstr);

                if (_appendTsOnTx)
                {
                    var txts = string.Format("\r[{0:HH:mm:ss.fff}] --> ", DateTime.Now);
                    AppendEventLog(txts, Color.DarkOliveGreen, false);
                }
                else
                {
                    AppendEventLog("\r", Color.DarkOliveGreen, false); //for richtextbox newline
                }

                var txbytes = new List<byte>();
                if (_txEncode == EncodeType.Auto)
                {
                    if (txstr.StartsWith("0x"))
                    {
                        txstr = txstr.Substring(2);
                        txbytes.AddRange(HexStringToByteArray(txstr));
                        if (_appendCrOnTx) txbytes.Add((byte)'\r');
                        if (_appendNlOnTx) txbytes.Add((byte)'\n');
                        _serialCom.Write(txbytes.ToArray());
                    }
                    else
                    {
                        if (_appendCrOnTx) txstr += "\r";
                        if (_appendNlOnTx) txstr += "\n";
                        _serialCom.Write(txstr);
                        txbytes.AddRange(Encoding.ASCII.GetBytes(txstr));
                    }
                    AppendEventLog(ByteArrayToFormatedString(txbytes.ToArray()), Color.DarkGreen, false);
                }
                else if (_txEncode == EncodeType.Hex)
                {
                    if (txstr.StartsWith("0x")) txstr = txstr.Substring(2);
                    txbytes.AddRange(HexStringToByteArray(txstr));
                    if (_appendCrOnTx) txbytes.Add((byte)'\r');
                    if (_appendNlOnTx) txbytes.Add((byte)'\n');
                    _serialCom.Write(txbytes.ToArray());
                    AppendEventLog(ByteArrayToHexString(txbytes.ToArray()), Color.DarkGreen, false);
                }
                else //ascii
                {
                    if (_appendCrOnTx) txstr += "\r";
                    if (_appendNlOnTx) txstr += "\n";
                    _serialCom.Write(txstr);
                    AppendEventLog(txstr, Color.DarkGreen, false);
                }

                richTextBoxTxData.Focus();
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        #endregion

        #region Date Received

        DateTime _lastRxTsTime = DateTime.Now;
        TimeSpan _appendTsOnRxInterval = TimeSpan.FromMilliseconds(20);

        private void SerialCom_OnDataReceived(object sender, SerialComEventArgs e)
        {
            try
            {
                var now = DateTime.Now;
                if (_appendTsOnRx && (now.Subtract(_lastRxTsTime) > _appendTsOnRxInterval))
                {
                    _lastRxTsTime = now;
                    var rxts = string.Format("\r[{0:HH:mm:ss.fff}] <-- ", _lastRxTsTime);
                    AppendEventLog(rxts, Color.MidnightBlue, false);
                }
                else if (_appendNlOnRx)
                {
                    AppendEventLog("\r", Color.MidnightBlue, false);  //for richtextbox newline
                }

                var rxstr = "";
                if (_rxDecode == EncodeType.Auto)
                {
                    rxstr = ByteArrayToFormatedString(e.Data);
                }
                else if (_rxDecode == EncodeType.Hex)
                {
                    rxstr = ByteArrayToHexString(e.Data);
                }
                else //ascii
                {
                    rxstr = Encoding.ASCII.GetString(e.Data);
                }
                AppendEventLog(rxstr, Color.Blue, false);
            }
            catch (Exception ex)
            {
                PopupException(ex.Message);
            }
        }

        private void SerialCom_OnException(object sender, SerialComEventArgs e)
        {
            AppendEventLog("Exception --> " + e.Message, Color.Magenta);
        }

        #endregion

    }
}
