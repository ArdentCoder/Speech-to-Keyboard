using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;

namespace speechRecog_Tutorial
{
    /// <summary>
    /// Interaction logic for textBoxToFile.xaml
    /// </summary>
    public partial class textBoxToFile : Window
    {
        public textBoxToFile()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            System.IO.File.WriteAllText("myfile.txt", textBox1.Text);
        }
    }
}
