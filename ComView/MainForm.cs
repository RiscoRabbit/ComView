using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;


namespace ComView
{
    public partial class MainForm : Form
    {
        static MainForm me;
        public MainForm()
        {
            InitializeComponent();
            me = this;
            devNames = new string[] { };
            check();
        }

        public static void SetText(string s)
        {
            if(me.InvokeRequired)
            {
                me.Invoke(new Action(() => SetText(s)), null);
                return;
            }
            me.textBox1.AppendText(s + Environment.NewLine);
        }
        string[] devNames;

        private void check()
        {
            string[] tmpDevNames = GetDeviceNames();
            if(tmpDevNames == null)
            {
                return;
            }
            foreach (string dev in tmpDevNames)
            {
                if (Array.IndexOf(devNames, dev) == -1)
                {
                    SetText("Add " + dev);
                }
            }
            foreach (string dev in devNames)
            {
                if (Array.IndexOf(tmpDevNames, dev) == -1)
                {
                    SetText("Del " + dev);
                }
            }
            devNames = tmpDevNames;
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            textBox1.Text = "";
            devNames = new string[] { };
            check();

        }
        public const int WM_DEVICECHANGE = 0x00000219;  //デバイス変化のWindowsイベントの値

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            switch (m.Msg)
            {
                case WM_DEVICECHANGE:   //デバイス状況の変化イベント
                    // AddMessage(m.ToString() + Environment.NewLine);
                    Task.Run(() => check());      //デバイスをチェック
                    break;
            }
        }


        // http://truthfullscore.hatenablog.com/entry/2014/01/10/180608
        public static string[] GetDeviceNames()
        {
            var deviceNameList = new System.Collections.ArrayList();
            var check = new System.Text.RegularExpressions.Regex("(COM[1-9][0-9]?[0-9]?)");
            ManagementClass mcPnPEntity = new ManagementClass("Win32_PnPEntity");
            ManagementObjectCollection manageObjCol = mcPnPEntity.GetInstances();

            foreach (ManagementObject manageObj in manageObjCol)
            {
                var namePropertyValue = manageObj.GetPropertyValue("Name");
                if (namePropertyValue == null)
                {
                    continue;
                }

                string name = namePropertyValue.ToString();
                if (check.IsMatch(name))
                {
                    deviceNameList.Add(name);
                }
            }

            if (deviceNameList.Count > 0)
            {
                string[] deviceNames = new string[deviceNameList.Count];
                int index = 0;
                foreach (var name in deviceNameList)
                {
                    deviceNames[index++] = name.ToString();
                }
                return deviceNames;
            }
            else
            {
                return null;
            }
        }
    }
}
