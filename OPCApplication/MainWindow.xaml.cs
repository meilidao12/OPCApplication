using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Services;
using CommunicationServers.OPC;
using OPCAutomation;
namespace OPCApplication
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        OPCClientHelper OPCClient = new OPCClientHelper();
        string ServerIP = "127.0.0.1";
        public MainWindow()
        {
            InitializeComponent();
            //this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            List<string> names = new List<string>();
            names = OPCClient.GetOPCServerNames(ServerIP);
            foreach(var name in names)
            {
                Console.WriteLine(name);
            }
            OPCServer opcServer = OPCClient.ConnectToServer(names[0], ServerIP);
            OPCBrowser opcbrowser = OPCClient.RecurBrowse(opcServer);
            foreach (var item in opcbrowser)
            {
                Console.WriteLine(item.ToString());
            }
            OPCClient.CreateGroup(opcServer,"tianchen");
            //OPCClient.GetServerInfo(opcServer);
            //Console.WriteLine(opcServer.OPCGroups.ToString());
        }
    }

   
}
