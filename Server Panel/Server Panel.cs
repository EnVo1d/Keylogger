using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;
using SimpleTcp;
using System.Diagnostics;

namespace Server_Panel
{
    public partial class ServerPanel : Form
    {
        private SimpleTcpServer socket;

        private readonly string dir = Directory.GetCurrentDirectory();
        private IConfiguration Configuration { get; set; }

        private bool isClicked = false;
        private int Port { get; set; }

        public ServerPanel()
        {
            InitializeComponent();
        }

        private void ServerPanel_Load(object sender, EventArgs e)
        {
            var builder = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            Configuration = builder.Build();
            var settings = this.Configuration.GetSection("AllowedPort").Get<Settings>();
            Port = settings.Port;
            if (!Directory.Exists(dir + "\\Logs"))
                Directory.CreateDirectory(dir + "\\Logs");
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (NetworkInterface.GetIsNetworkAvailable())
            {
                if (!isClicked)
                {
                    socket = new SimpleTcpServer(GetLocalIPAddress(), Port);
                    socket.Settings.StreamBufferSize = 1024;
                    socket.Events.ClientConnected += Events_ClientConnected;
                    socket.Events.ClientDisconnected += Events_ClientDisconnected;
                    socket.Events.DataReceived += Events_DataReceived;
                    socket.StartAsync();
                    listBox2.Items.Add($"Server started{Environment.NewLine}");
                    isClicked = !isClicked;
                    btnStart.Text = "Stop";
                }
                else
                {
                    isClicked = !isClicked;
                    btnStart.Text = "Start";
                    socket.Stop();
                    socket.Dispose();
                    listBox2.Items.Add($"Server stoped{Environment.NewLine}");
                }
            }
            else listBox2.Items.Add("Network is unavailable");
        }

        private void Events_DataReceived(object sender, SimpleTcp.DataReceivedEventArgs e)
        {
            string data = Encoding.UTF8.GetString(e.Data);
            Debug.WriteLine(DateTime.Now.ToString() + $"{data}");
            try
            {
                using (FileStream fs = new FileStream(dir + "/Logs" + $"/{e.IpPort.Substring(0, e.IpPort.IndexOf(":"))}/" + DateTime.Now.Date.ToLongDateString() + ".txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    fs.Write(e.Data, 0, e.Data.Length);
                }
            }
            catch { }
        }

        private void Events_ClientDisconnected(object sender, ClientDisconnectedEventArgs e)
        {
            this.Invoke((MethodInvoker)(() =>
            {
                listBox2.Items.Add($"[{e.IpPort}] disconnected: {e.Reason}{Environment.NewLine}");
                listBox1.Items.Remove(e.IpPort);
            }));
        }

        private void Events_ClientConnected(object sender, ClientConnectedEventArgs e)
        {
            this.Invoke((MethodInvoker)(() =>
            {
                listBox2.Items.Add($"[{e.IpPort}] connected{Environment.NewLine}");
                listBox1.Items.Add(e.IpPort);
            }));
            Directory.CreateDirectory(dir + "/Logs" + $"/{e.IpPort.Substring(0, e.IpPort.IndexOf(":"))}");
        }

        private string GetLocalIPAddress()
        {
            var Host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in Host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            return null;
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                int index = listBox1.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    string clickedItem = listBox1.Items[index].ToString();
                    Process.Start("explorer.exe", dir +"\\Logs"+"\\"+clickedItem.Substring(0, clickedItem.IndexOf(":")));
                }
            }
        }
    }
}
