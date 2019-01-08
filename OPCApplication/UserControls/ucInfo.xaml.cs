using OPCApplication.Models;
using Services.DataBase;
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
using System.Windows.Threading;

namespace OPCApplication.UserControls
{
    /// <summary>
    /// ucInfo.xaml 的交互逻辑
    /// </summary>
    public partial class ucInfo : UserControl
    {
        AccessHelper access = new AccessHelper();
        DispatcherTimer ClockTimer = new DispatcherTimer(); //采集plc数据
        public ucInfo()
        {
            InitializeComponent();
            this.Loaded += UcInfo_Loaded;
        }

        private void UcInfo_Loaded(object sender, RoutedEventArgs e)
        {
            this.ClockTimer.Interval = TimeSpan.FromSeconds(1);
            this.ClockTimer.Tick += ClockTimer_Tick;
            this.Refresh();
        }

        private void ClockTimer_Tick(object sender, EventArgs e)
        {
            List<OpcListModel> models = access.GetDataTable<OpcListModel>("Select * from [OPCConfig] order by Id asc");
            var model = models.Find(m => m.Name == this.txbTags.Text);
            if (model != null)
            {
                this.txbAddress.Text = model.Address;
                this.txbValue.Text = model.Value;
            }
        }
        public string Opcname
        {
            set
            {
                this.txbTags.Text = value;
            }
        }

        public void Refresh()
        {
            if (!string.IsNullOrEmpty(this.txbTags.Text))
            {
                List<OpcListModel> models = access.GetDataTable<OpcListModel>("Select * from [OPCConfig] order by Id asc");
                var model = models.Find(m => m.Name == this.txbTags.Text);
                if (model != null)
                {
                    this.txbAddress.Text = model.Address;
                    this.txbValue.Text = model.Value;
                    this.txbInfo.Text = model.Info;
                    if(!this.ClockTimer.IsEnabled)
                    {
                        this.ClockTimer.Start();
                    }
                }
                else
                {
                    this.ClockTimer.Stop();
                }
            }
            else
            {
                this.ClockTimer.Stop();
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string commandtext = string.Format("UPDATE [OPCConfig] SET [info]='{0}' WHERE Name='{1}'", this.txbInfo.Text, this.txbTags.Text);
            if (!access.Execute(commandtext))
            {
                MessageBox.Show("保存失败");
            }
        }
    }
}
