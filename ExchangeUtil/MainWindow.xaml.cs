using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace ExchangeUtil
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<ExchangeHelper.ExchangeData> GridItems;

        List<string> _lstDeleteItems;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LblInfo.Content = "Getting '" + TxtUser.Text + "' appointment's .......";

            if (string.IsNullOrEmpty(TxtLogin.Text))
                MessageBox.Show("Login id is empty");

            if (string.IsNullOrEmpty(TxtPassword.Text))
                MessageBox.Show("Password is empty");

            if (string.IsNullOrEmpty(TxtUrl.Text))
                MessageBox.Show("Url is empty");

            if (string.IsNullOrEmpty(TxtUser.Text))
                MessageBox.Show("User is empty");

            GetUsersData();
        }

        private void chkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (ExchangeHelper.ExchangeData c in DgData.ItemsSource)
            {
                c.IsSelected = true;
            }

            DgData.Items.Refresh();
        }

        private void chkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (ExchangeHelper.ExchangeData c in DgData.ItemsSource)
            {
                c.IsSelected = false;
            }

            DgData.Items.Refresh();
            
        }

        /// <summary>
        /// 
        /// </summary>
        private void GetUsersData()
        {
            _lstDeleteItems = new List<string>();
            ExchangeHelper.Login = TxtLogin.Text;
            ExchangeHelper.Password = TxtPassword.Text;
            ExchangeHelper.Url = TxtUrl.Text;
            ExchangeHelper.User = TxtUser.Text;
            int.TryParse(TxtBackDays.Text, out int backDays);
            ExchangeHelper.BackDays = -backDays;
            GridItems = ExchangeHelper.GetUserAppointments();
            LblInfo.Content = "Total records : " + GridItems?.Count;
            DgData.ItemsSource = GridItems;
            if (GridItems?.Count > 1)
                BtnDelete.IsEnabled = true;
        }



        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            _lstDeleteItems = new List<string>();
            foreach (var item in DgData.Items)
            {
                if(item is ExchangeHelper.ExchangeData exchangeData && exchangeData.IsSelected)
                    _lstDeleteItems.Add(exchangeData.UniqueId);
            }

            if (_lstDeleteItems.Any())
            {
                ExchangeHelper.DeleteItems(_lstDeleteItems);
                LblInfo.Content = "Deleted rows : " + _lstDeleteItems?.Count;
                GridItems = ExchangeHelper.GetUserAppointments();
                DgData.ItemsSource = GridItems;
            }
        }

       
    }
}
