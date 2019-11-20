using FirmaPdfWpf.ViewModels.Pin;
using MahApps.Metro.Controls;
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
using System.Windows.Shapes;

namespace FirmaPdfWpf
{
    /// <summary>
    /// Interaction logic for pin.xaml
    /// </summary>
    public partial class Pin : MetroWindow
    {
        PinViewModel model;

        public Pin(PinViewModel model)
        {
            InitializeComponent();
            this.model = model;
            this.DataContext = model;
            txtPin.Focus();

        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            model.Pin = txtPin.Password;
            this.DialogResult = true;
            this.Close();
        }
    }
}
