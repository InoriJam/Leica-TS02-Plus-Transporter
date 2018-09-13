using System.Windows;
using System.IO;
using System.IO.Ports;
using System.Collections;
using System.Timers;
using Microsoft.Win32;
using System.Text;
using System;

namespace DMSkin.WPF.Demos
{
    public partial class StartWindow 
    {
        class TextAndValue<T>
        {
            private T _RealValue;
            private string _DisplayText = "";

            public string DisplayText
            {
                get
                {
                    return _DisplayText;
                }
            }

            public T RealValue
            {
                get
                {
                    return _RealValue;
                }
            }

            public TextAndValue(string ShowText, T RealVal)
            {
                _DisplayText = ShowText;
                _RealValue = RealVal;
            }

            public override string ToString()
            {
                return base.ToString();
            }

        }

        System.Timers.Timer timer = new System.Timers.Timer();
        System.Timers.Timer timer2 = new System.Timers.Timer();
        SerialPort serialPort1;

        public StartWindow()
        {
            InitializeComponent();
            //初始化端口选择
            foreach(string com in System.IO.Ports.SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(com);
            }
            comboBox1.SelectedIndex = 0;

            //初始化波特率选择
            ArrayList combo2_data = new ArrayList();
            combo2_data.Add(new TextAndValue<int>("1200", 1200));
            combo2_data.Add(new TextAndValue<int>("2400", 2400));
            combo2_data.Add(new TextAndValue<int>("4800", 4800));
            combo2_data.Add(new TextAndValue<int>("9600", 9600));
            combo2_data.Add(new TextAndValue<int>("14400", 14400));
            combo2_data.Add(new TextAndValue<int>("19200", 19200));
            combo2_data.Add(new TextAndValue<int>("57600", 57600));
            combo2_data.Add(new TextAndValue<int>("115200", 115200));
            comboBox2.ItemsSource = combo2_data;
            comboBox2.DisplayMemberPath = "DisplayText";
            comboBox2.SelectedValuePath = "RealValue";
            comboBox2.SelectedIndex = 3;


            //初始化奇偶校验选择
            ArrayList combo3_data = new ArrayList();
            combo3_data.Add(new TextAndValue<System.IO.Ports.Parity>("无(N)", System.IO.Ports.Parity.None));
            combo3_data.Add(new TextAndValue<System.IO.Ports.Parity>("奇(O)", System.IO.Ports.Parity.Odd));
            combo3_data.Add(new TextAndValue<System.IO.Ports.Parity>("偶(E)", System.IO.Ports.Parity.Even));
            comboBox3.ItemsSource = combo3_data;
            comboBox3.DisplayMemberPath = "DisplayText";
            comboBox3.SelectedValuePath = "RealValue";
            comboBox3.SelectedIndex = 2;

            //判断是否接受完成的计时器
            timer.Elapsed += new ElapsedEventHandler(end);
            timer.Interval = 7000;
            timer.AutoReset = false;

            //判断是否通信超时的计时器
            timer2.Elapsed += new ElapsedEventHandler(timeout);
            timer2.Interval = 8000;
            timer2.AutoReset = false;

            //初始化radiobutton
            radioButton3.IsChecked = true;

            //初始化串口
            serialPort1 = new SerialPort();
            serialPort1.DataReceived += serialPort1_DataReceived;
        }

        StreamWriter wt;
        private void DMButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            if ((bool)radioButton1.IsChecked)
            {
                dialog.Filter = "XML(*.xml)|*.xml";
            }
            else if ((bool)radioButton2.IsChecked)
            {
                dialog.Filter = "DXF(*.dxf)|*.dxf";
            }
            else if ((bool)radioButton3.IsChecked)
            {
                dialog.Filter = "CASS(*.dat)|*.dat";
            }
            else if ((bool)radioButton4.IsChecked)
            {
                dialog.Filter = "GSI(*.gsi)|*.gsi";
            }
            else if ((bool)radioButton5.IsChecked)
            {
                dialog.Filter = "SV(*.csv)|*.csv";
            }
            dialog.ShowDialog();
            if (dialog.FileName == "")
            {
                return;
            }
            FileStream fs = new FileStream(dialog.FileName, FileMode.Create);
            wt = new StreamWriter(fs, Encoding.ASCII);
            serialPort1.BaudRate = (comboBox2.SelectedItem as TextAndValue<int>).RealValue;
            serialPort1.PortName = comboBox1.SelectedItem as string;
            serialPort1.ReceivedBytesThreshold = 1;
            serialPort1.DataBits = 8;
            serialPort1.Parity = (comboBox3.SelectedItem as TextAndValue<System.IO.Ports.Parity>).RealValue;
            serialPort1.StopBits = System.IO.Ports.StopBits.One;
            serialPort1.ReadTimeout = 5000;
            try
            {
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("端口打开失败，请检查驱动程序");
                return;
            }
            MessageBox.Show("请点击确定后在全站仪上点击导出");
            timer2.Start();
        }

        private void end(object source, System.Timers.ElapsedEventArgs e)
        {
            wt.Close();
            serialPort1.Close();
            MessageBox.Show("传输完成");
        }
        private void timeout(object source, System.Timers.ElapsedEventArgs e)
        {
            wt.Close();
            serialPort1.Close();
            MessageBox.Show("通信超时，请及时在全站仪上点击发送或者检查是否选择了设备对应的端口");
            return;
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            timer2.Stop();
            try
            {
                if (serialPort1.BytesToRead != 0)
                {
                    timer.Stop();
                    wt.WriteLine(serialPort1.ReadLine());
                    timer.Start();
                }
            }
            catch (TimeoutException ex)
            {
                wt.Close();
                serialPort1.Close();
                MessageBox.Show("读取超时");
            }
        }

    }
}
