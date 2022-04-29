using System;
using System.Windows;
using System.Windows.Controls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Input;

namespace Laba_26_KPIAP
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SqlConnection sqlConnection = null;
        string connectionString;
        DataTable Books;
        public MainWindow()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["Books"].ConnectionString;
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            UpdateDB();
        }
        private void UpdateDB()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

            sqlConnection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM [Books]", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            List<Name> Equipment = new List<Name>();
            while (reader.Read())
            {
                int id = Convert.ToInt32(reader["ID"]);
                string nname = Convert.ToString(reader["Name"]);
                string author = Convert.ToString(reader["Author"]);
                string cover = Convert.ToString(reader["Cover"]);
                string pages = Convert.ToString(reader["Pages"]);
                string cost = Convert.ToString(reader["Cost"]);
                Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
            }
            moviesGrid.ItemsSource = Equipment;
        }
        private void UpdateDB(object sender, RoutedEventArgs e)
        {
            UpdateDB();
        }

        private int Findid()
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [Books]", connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int ID = Convert.ToInt32(reader["ID"]);
                string Name = Convert.ToString(reader["Name"]);

                if (Name == Deletebox.Text)
                {
                    return ID;
                }
            }
            return 0;
        }
        private void DeleteDB(object sender, RoutedEventArgs e)
        {

            string sql = $"DELETE FROM [Books] WHERE Id={Findid()}";
            Books = new DataTable();
            SqlConnection connection = null;
            connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand(sql, connection);

            connection.Open();
            command.ExecuteNonQuery();
            connection.Close();

            UpdateDB();
        }

        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int.Parse(PagesBox.Text);
                int.Parse(CostBox.Text);
                if (OuterBox.Text == "Твердый" || OuterBox.Text == "Мягкий")
                {
                    if (Convert.ToInt32(CostBox.Text) > 0 && Convert.ToInt32(PagesBox.Text) > 0)
                    {
                        string sql = $"INSERT INTO [Books] (Name, Author, Cover, Pages, Cost) VALUES (N'{NameBox.Text}', N'{AuthorBox.Text}', N'{OuterBox.Text}', N'{PagesBox.Text}', N'{CostBox.Text}')";
                        SqlConnection connection = connection = new SqlConnection(connectionString);
                        connection.Open();

                        SqlCommand command = new SqlCommand(sql, connection);
                        command.ExecuteNonQuery();

                        connection.Close();
                    }
                    else
                    {
                        MessageBox.Show("Неправильно введены страницы или цена", "Ошибка");
                    }

                }
                else if (OuterBox.Text == "Твердый/Мягкий")
                {
                    if (Convert.ToInt32(CostBox.Text) > 0 && Convert.ToInt32(PagesBox.Text) > 0)
                    {
                        string sql = $"INSERT INTO [Books] (Name, Author, Cover, Pages, Cost) VALUES (N'{NameBox.Text}', N'{AuthorBox.Text}', N'{OuterBox.Text}', N'{PagesBox.Text}', N'{CostBox.Text}')";
                        SqlConnection connection = connection = new SqlConnection(connectionString);
                        connection.Open();

                        SqlCommand command = new SqlCommand(sql, connection);
                        command.ExecuteNonQuery();

                        connection.Close();
                    }
                    else
                    {
                        MessageBox.Show("Неправильно введены страницы или цена", "Ошибка");
                    }
                }
                else
                {
                    MessageBox.Show("Неправильно введены переплет", "Ошибка");
                }

            }
            catch (FormatException ex)
            {
                MessageBox.Show("Неправильно введены данные", "Ошибка");
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Полностью заполните поля", "Ошибка");
            }




            UpdateDB();
        }

        private void TwoParametarsSort_Click(object sender, RoutedEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

            sqlConnection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Name, Cost", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            List<Name> Equipment = new List<Name>();
            while (reader.Read())
            {
                int id = Convert.ToInt32(reader["ID"]);
                string nname = Convert.ToString(reader["Name"]);
                string author = Convert.ToString(reader["Author"]);
                string cover = Convert.ToString(reader["Cover"]);
                string pages = Convert.ToString(reader["Pages"]);
                string cost = Convert.ToString(reader["Cost"]);
                Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
            }
            moviesGrid.ItemsSource = Equipment;
        }

        private void HardLightCover_Click(object sender, RoutedEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

            sqlConnection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM [Books]", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            List<Name> Equipment = new List<Name>();
            while (reader.Read())
            {
                int id = Convert.ToInt32(reader["ID"]);
                string nname = Convert.ToString(reader["Name"]);
                string author = Convert.ToString(reader["Author"]);
                string cover = Convert.ToString(reader["Cover"]);
                string pages = Convert.ToString(reader["Pages"]);
                string cost = Convert.ToString(reader["Cost"]);
                if (cover == "Твердый/Мягкий")
                {
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
            }
            moviesGrid.ItemsSource = Equipment;
        }

        private void BooksMoreThan10_Click(object sender, RoutedEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

            sqlConnection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM [Books]", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            List<Name> Equipment = new List<Name>();
            while (reader.Read())
            {
                int id = Convert.ToInt32(reader["ID"]);
                string nname = Convert.ToString(reader["Name"]);
                string author = Convert.ToString(reader["Author"]);
                string cover = Convert.ToString(reader["Cover"]);
                string pages = Convert.ToString(reader["Pages"]);
                string cost = Convert.ToString(reader["Cost"]);
                if (Convert.ToInt32(cost) > 10 && cover == "Мягкий")
                {
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
            }
            moviesGrid.ItemsSource = Equipment;
        }

        private void MaxPages_Click(object sender, RoutedEventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

            sqlConnection.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Pages DESC", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            List<Name> Equipment = new List<Name>();
            int max = 0;

            while (reader.Read())
            {

                int id = Convert.ToInt32(reader["ID"]);
                string nname = Convert.ToString(reader["Name"]);
                string author = Convert.ToString(reader["Author"]);
                string cover = Convert.ToString(reader["Cover"]);
                string pages = Convert.ToString(reader["Pages"]);
                string cost = Convert.ToString(reader["Cost"]);
                Name maxN = new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost };
                if (max < Convert.ToInt32(pages))
                {
                    max = Convert.ToInt32(pages);
                    Equipment.Add(maxN);
                }
            }
            moviesGrid.ItemsSource = Equipment;
        }
        private void GroupColumns_Click(object sender, RoutedEventArgs e)
        {
            if (RadioButton1.IsChecked == true)
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

                sqlConnection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Name", sqlConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<Name> Equipment = new List<Name>();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    string nname = Convert.ToString(reader["Name"]);
                    string author = Convert.ToString(reader["Author"]);
                    string cover = Convert.ToString(reader["Cover"]);
                    string pages = Convert.ToString(reader["Pages"]);
                    string cost = Convert.ToString(reader["Cost"]);
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
                moviesGrid.ItemsSource = Equipment;
            }
            else if (RadioButton2.IsChecked == true)
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

                sqlConnection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Author", sqlConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<Name> Equipment = new List<Name>();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    string nname = Convert.ToString(reader["Name"]);
                    string author = Convert.ToString(reader["Author"]);
                    string cover = Convert.ToString(reader["Cover"]);
                    string pages = Convert.ToString(reader["Pages"]);
                    string cost = Convert.ToString(reader["Cost"]);
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
                moviesGrid.ItemsSource = Equipment;
            }
            else if (RadioButton3.IsChecked == true)
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

                sqlConnection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Cover", sqlConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<Name> Equipment = new List<Name>();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    string nname = Convert.ToString(reader["Name"]);
                    string author = Convert.ToString(reader["Author"]);
                    string cover = Convert.ToString(reader["Cover"]);
                    string pages = Convert.ToString(reader["Pages"]);
                    string cost = Convert.ToString(reader["Cost"]);
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
                moviesGrid.ItemsSource = Equipment;
            }
            else if (RadioButton4.IsChecked == true)
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

                sqlConnection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Pages", sqlConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<Name> Equipment = new List<Name>();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    string nname = Convert.ToString(reader["Name"]);
                    string author = Convert.ToString(reader["Author"]);
                    string cover = Convert.ToString(reader["Cover"]);
                    string pages = Convert.ToString(reader["Pages"]);
                    string cost = Convert.ToString(reader["Cost"]);
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
                moviesGrid.ItemsSource = Equipment;
            }
            else if (RadioButton5.IsChecked == true)
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Books"].ConnectionString);

                sqlConnection.Open();
                SqlCommand command = new SqlCommand("SELECT * FROM [Books] ORDER BY Cost", sqlConnection);
                SqlDataReader reader = command.ExecuteReader();
                List<Name> Equipment = new List<Name>();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    string nname = Convert.ToString(reader["Name"]);
                    string author = Convert.ToString(reader["Author"]);
                    string cover = Convert.ToString(reader["Cover"]);
                    string pages = Convert.ToString(reader["Pages"]);
                    string cost = Convert.ToString(reader["Cost"]);
                    Equipment.Add(new Name() { ID = id, Название = nname, Автор = author, Обложка = cover, КОЛСтраниц = pages, Цена = cost });
                }
                moviesGrid.ItemsSource = Equipment;
            }
            else
            {
                MessageBox.Show("Выберите по какому полю будет идти группировка", "Ошибка");
            }
        }
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            moviesGrid.SelectAllCells();
            moviesGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, moviesGrid);
            moviesGrid.UnselectAllCells();
            var result = (string)Clipboard.GetData(DataFormats.Text);
            dynamic wordApp = null;
            try
            {
                var sw = new StreamWriter("export.doc");
                sw.WriteLine(result);
                sw.Close();
                //var proc = Process.Start("export.doc");
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                wordApp = Activator.CreateInstance(wordType);
                wordApp.Documents.Add(System.AppDomain.CurrentDomain.BaseDirectory + "export.doc");
                wordApp.ActiveDocument.Range.ConvertToTable(1, moviesGrid.Items.Count, moviesGrid.Columns.Count);
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
                // ignored
            }
        }
        public class Name
        {
            public int ID { get; set; }
            public string Название { get; set; }
            public string Автор { get; set; }
            public string Обложка { get; set; }
            public string КОЛСтраниц { get; set; }
            public string Цена { get; set; }
        }

       
    }
}
