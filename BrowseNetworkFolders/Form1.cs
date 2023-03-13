using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BrowseNetworkFolders
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                RegistryKey localMachine = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64); //here you specify where exactly you want your entry

                var reg = localMachine.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System", true);
                if (reg == null)
                {
                    reg = localMachine.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System", true);
                }

                if (reg.GetValue("EnableLinkedConnections") == null)
                {
                    reg.SetValue("EnableLinkedConnections", "1", RegistryValueKind.DWord);
                    MessageBox.Show(
                        "Your configuration is now created,you have to restart your device for settings to take effect");
                }
                object oo =reg.GetValue("EnableLinkedConnections");
                MessageBox.Show("EnableLinkedConnections:"+oo.ToString());

                MessageBox.Show("complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please right click and select 'Run as admin'" + ex.Message);
            }
        }
    }
}
