﻿using Microsoft.Win32;
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
using System.IO;

namespace GenTemplateBJ
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ExcelConverters converter = new();
        public MainWindow()
        {
            InitializeComponent();
            DataContext=converter;
        }

        private void Input_Click(object sender, RoutedEventArgs e)
        {
            var excel = Utils.OpenAnExcelFile();
            if (excel is not null)
            {
                converter.ExcelData = new InputExcelData(excel);
            }
            // MessageBox.Show(converter.IsExcelDataNotNull.ToString());
        }

        private void Output_Click(object sender, RoutedEventArgs e)
        {
            var path = Utils.OpenAFolder();
            if (path is not null)
            {
                foreach((var key, var value) in converter.OutputExcels) 
                {
                    value.Workbook.SaveAs(Path.Combine(path, key));
                }
                foreach((var key, var value) in converter.OutputDocxs)
                {
                    using var stream = new FileStream(Path.Combine(path, key), FileMode.Create);
                    value.Write(stream);
                }
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            var templateType = TemplateType.Text;
            if(!converter.TemplateTypeToExcelConverter.ContainsKey(templateType)) 
            {
                _ = MessageBox.Show("请选择有效的模板类型。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            converter.TemplateTypeToExcelConverter[templateType]();
        }
    }
}
