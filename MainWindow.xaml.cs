using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Aspose.Cells;
using System.IO;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows.Controls;

namespace VPNGen
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();            
        }

        private int numberOfPasswords { get; set; }
        private string loginPrefiks { get; set; }
        private int startingNumber { get; set; }
        private string profile { get; set; }
        private string service { get; set; }
        private string path { get; set; }


        const string Digits = "0123456789";
        const string Alphabet = "abcdefghijklmnopqrstuvwxyz";
        const string Symbols = " ~`@#$%^&*()_+-=[]{};'\\:\"|,./<>?";

        private void TekeData()
        {
            numberOfPasswords = int.Parse(textBox1.Text);
            loginPrefiks = textBox2.Text;
            startingNumber = int.Parse(textBox3.Text);
            profile = textBox4.Text;
            service = textBox5.Text;                        
        }

        private void Button_Click(object sender, EventArgs e)
        {
            try
            {
                TekeData();
                TakePath();
                MainLogic();
                MessageBox.Show("Сгенерировано!");
            }
            catch (Exception)
            {
                MessageBox.Show("Не коректный ввод!");
            }          
                        
        }

        private void TakePath()
        {
            using CommonOpenFileDialog dlg = new CommonOpenFileDialog("Выберете папку");
            dlg.IsFolderPicker = true;
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dlg.Multiselect = false;
            dlg.AllowNonFileSystemItems = false;
            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {

                path = dlg.FileName;                
            }           
        }
       



        [Flags]
        private enum PasswordChars
        {
            Digits = 0b0001,
            Alphabet = 0b0010,
            Symbols = 0b0100
        }

        private void MainLogic()
        {         

            List<string> listPass = new List<string>();
            listPass.Add("/ppp secret");

            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            int passwordLength = 16;
            int charSet = 3;            

            for (int i = 1; i < numberOfPasswords + 1; i++)
            {
                string genpass = GeneratePassword((PasswordChars)charSet, passwordLength);
                string psassword = string.Format("add name={0}{1} password={2} profile={3} service={4}",
                    loginPrefiks, startingNumber, genpass, profile, service);
                listPass.Add(psassword);

                Cell cell_A = sheet.Cells["A" + i];
                Cell cell_B = sheet.Cells["B" + i];

                cell_A.PutValue(loginPrefiks + startingNumber);
                cell_B.PutValue(genpass);

                startingNumber++;
            }

            string txtPath = System.IO.Path.Combine(path, string.Format(@"Пароли_{0}.txt", loginPrefiks));
            string xlsxPath = System.IO.Path.Combine(path, string.Format(@"Пароли_{0}.xlsx", loginPrefiks));

            File.AppendAllLines(txtPath, listPass.Select(t => t.ToString()));
            wb.Save(xlsxPath, SaveFormat.Xlsx);            

        }

        private string GeneratePassword(PasswordChars passwordChars, int length)
        {
            var random = new Random();
            var resultPassword = new StringBuilder(length);
            var passwordCharSet = string.Empty;
            if (passwordChars.HasFlag(PasswordChars.Alphabet))
            {
                passwordCharSet += Alphabet + Alphabet.ToUpper();
            }
            if (passwordChars.HasFlag(PasswordChars.Digits))
            {
                passwordCharSet += Digits;
            }
            if (passwordChars.HasFlag(PasswordChars.Symbols))
            {
                passwordCharSet += Symbols;
            }
            for (var i = 0; i < length; i++)
            {
                resultPassword.Append(passwordCharSet[random.Next(0, passwordCharSet.Length)]);
            }
            return resultPassword.ToString();
        }
    }


    

}
