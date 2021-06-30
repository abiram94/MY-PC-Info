using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Net;
using System.Net.Sockets;
using System.Management;

namespace Mypc_Info
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblcompname.Text = System.Environment.GetEnvironmentVariable("COMPUTERNAME"); // Computername
            lblusername.Text = System.DirectoryServices.AccountManagement.UserPrincipal.Current.UserPrincipalName; // Username
            lblipaddress.Text = GetLocalIPAddress();
            lblmac.Text = GetMACAddress();
            lblsno.Text = getosname();
            lblmotherno.Text = getmothername();
            lblramsize.Text = getramsize();
            lblofficeversion.Text = getofficeversion();
        }


        public string GetMACAddress()
        {
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            String sMacAddress = string.Empty;
            foreach (NetworkInterface adapter in nics)
            {
                if (sMacAddress == String.Empty)// only return MAC Address from first card
                {
                    IPInterfaceProperties properties = adapter.GetIPProperties();
                    sMacAddress = adapter.GetPhysicalAddress().ToString();
                }
            } return sMacAddress;
        }

        public static string GetLocalIPAddress()
        {
            var host = System.Net.Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }

        public static string getosname()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
            foreach (ManagementObject os in searcher.Get())
            {
                result = os["Caption"].ToString();
                break;
            }
            return result;
        }

        public static string getmothername()
        {
            string result2 = "";
            string result0 = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Product, SerialNumber,Manufacturer FROM Win32_BaseBoard");
            foreach (ManagementObject os in searcher.Get())
            {
                result2 = os["Product"].ToString();
                result0 = os["Manufacturer"].ToString();
                break;
            }
            return result0 + "-" + result2;
        }

        public static string getramsize()
        {
            string ramsize = "";
            string Query = "SELECT Capacity FROM Win32_PhysicalMemory";

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(Query);

            UInt64 Capacity = 0;
            foreach (ManagementObject WniPART in searcher.Get())
            {
                Capacity += (Convert.ToUInt64(WniPART.Properties["Capacity"].Value) / 1048576 / 1024);
                ramsize = Capacity.ToString() + " GB";
            }

            return ramsize;
        }

        public static string getofficeversion()
        {
            string sVersion = string.Empty;
            Microsoft.Office.Interop.Word.Application appVersion = new Microsoft.Office.Interop.Word.Application();
            appVersion.Visible = false;
            switch (appVersion.Version.ToString())
            {
                case "7.0":
                    sVersion = "95";
                    break;
                case "8.0":
                    sVersion = "97";
                    break;
                case "9.0":
                    sVersion = "2000";
                    break;
                case "10.0":
                    sVersion = "2002";
                    break;
                case "11.0":
                    sVersion = "2003";
                    break;
                case "12.0":
                    sVersion = "2007";
                    break;
                case "14.0":
                    sVersion = "2010";
                    break;
                case "15.0":
                    sVersion = "2013";
                    break;
                case "16.0":
                    sVersion = "2016 OR Above";
                    break;
                default:
                    sVersion = "OOPS";
                    break;
            }
            return sVersion;
        }


    }
}
