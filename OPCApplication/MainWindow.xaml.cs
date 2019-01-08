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
using Services.DataBase;
using System.Data;
using OPCApplication.Models;
namespace OPCApplication
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        SqlHelper sql = new SqlHelper();
        AccessHelper access = new AccessHelper();
        OPCClientHelper OPCClient = new OPCClientHelper();
        string ServerIP = "127.0.0.1";
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            IniOpc();
            AddOpcTags();
        }

        /// <summary>
        /// opc初始化
        /// </summary>
        private void IniOpc()
        {
            List<string> names = new List<string>();
            OPCClient.DataChangeEvent += OPCClient_DataChangeEvent;
            names = OPCClient.GetOPCServerNames(ServerIP);
            OPCServer opcServer = OPCClient.ConnectToServer(names[0], ServerIP);
            OPCBrowser opcbrowser = OPCClient.RecurBrowse(opcServer);
            foreach (var item in opcbrowser)
            {
                this.lstTags.Items.Add(item.ToString());
            }
            OPCClient.CreateGroup(opcServer, "tianchen");
        }

        /// <summary>
        /// opc添加标签
        /// </summary>
        private void AddOpcTags()
        {
            List<OpcListModel> models = access.GetDataTable<OpcListModel>("Select * from [OPCConfig] order by Id asc");
            foreach(var item in models)
            {
                this.lstAddedTags.Items.Add(item.Name);
                OPCClient.AddItem(item.Name);
            }
        }

        private void OPCClient_DataChangeEvent(List<OPCDataItem> OpcDataItems)
        {
            foreach(var item in OpcDataItems)
            {
                string commandtext = string.Format("UPDATE [OPCConfig] SET [Value]={0} WHERE Name='{1}'", item.ItemValue, item.ItemName);
                access.Execute(commandtext);
            }
        }



        private void lstTags_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string selectedText = (sender as ListBox).SelectedItem.ToString();
            string commandtext = string.Format("INSERT INTO [OPCConfig] (Name) values ('{0}')", selectedText);
            if (access.Execute(commandtext))
            {
                this.lstAddedTags.Items.Add(selectedText);
                OPCClient.AddItem(selectedText);
            }
        }

        private void btnAddressSerial_Click(object sender, RoutedEventArgs e)
        {
            List<OpcListModel> models = access.GetDataTable<OpcListModel>("Select * from [OPCConfig] order by Id asc");
            int addr = 40001;
            foreach (var item in models)
            {
                string commandtext = string.Format("UPDATE [OPCConfig] SET [Address]={0} WHERE Name='{1}'", addr.ToString(), item.Name);
                if(access.Execute(commandtext))
                {
                    addr += 2;
                }
                else
                {
                    MessageBox.Show("序列化失败");
                    return;
                }
            }
            MessageBox.Show("序列化成功");
        }

        private void lstAddedTags_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedText = (sender as ListBox).SelectedItem.ToString();
            this.opcInfo.Opcname = selectedText;
            this.opcInfo.Refresh();
        }
    }
}
