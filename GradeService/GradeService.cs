using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.ServiceProcess;
using System.Text.RegularExpressions;

namespace GradeService
{
    public partial class GradeService : ServiceBase
    {

        public GradeService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            StartFileSystemWatcher();
        }

        protected override void OnStop()
        {
        }

        //  Bu metot bir dosya oluştuğu zaman çağrılır
        private static void OnCreated(object source, FileSystemEventArgs e)
        {
            // Grades klasöründe dosya oluşma işlemi tamamlandıktan sonra işlemler yapılsın
            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                // Oluşturulan ya da eklenen .cvs dosyası burada SQLite veritabanına kaydedilecek
                string csvPath = e.FullPath;

                // ConvertCSVtoDataTable fonsiyonu ile DataTable oluştur
                DataTable dt = ConvertCSVtoDataTable(csvPath);

                // DataTable'ı SQLite veritabanı dosyasına kaydet
                string dbName = Path.GetFileNameWithoutExtension(e.FullPath);
                string dbPath = @"C:\Grades\" + dbName + ".db";
                if (!File.Exists(dbPath))
                    SQLiteConnection.CreateFile(dbPath);
                // Eğer aynı isimli db varsa oluşacak veritabanı için yeni bir ad generate et ve yeni adı ile dosya oluştur
                else
                {
                    int dbNameGenerator = 1;
                    while (File.Exists(dbPath))
                    {
                        dbName = Path.GetFileNameWithoutExtension(e.FullPath);
                        dbName = string.Format(dbName + " ({0})", dbNameGenerator.ToString());
                        dbPath = @"C:\Grades\" + dbName + ".db";
                        dbNameGenerator++;
                    }
                    SQLiteConnection.CreateFile(dbPath);
                }

                using (SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source=" + dbPath + ";Version=3;"))
                {
                    m_dbConnection.Open();
                    string tableName = Regex.Split(dbName, @"\s+").ToString();
                    int affectedRows = -1;
                    using (SQLiteCommand m_dbCommand = new SQLiteCommand(m_dbConnection))
                    {
                        m_dbCommand.CommandText = $"CREATE TABLE IF NOT EXISTS '{tableName}' (ders varchar(20), ogrenci_no varchar(20), vize1 int, vize2 int, vize3 int, final int)";
                        m_dbCommand.ExecuteNonQuery();
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        using (SQLiteCommand m_dbCommand = new SQLiteCommand(m_dbConnection))
                        {
                            // Ders kodu ve öğrenci numarası aynı olan kayıttan bir tane olması gerek
                            m_dbCommand.CommandText = $"SELECT COUNT(ders) as sayi FROM '{tableName}' WHERE ders = @ders AND ogrenci_no = @ogrenci_no";
                            m_dbCommand.Parameters.AddWithValue("@ders", dt.Rows[i][0].ToString());
                            m_dbCommand.Parameters.AddWithValue("@ogrenci_no", dt.Rows[i][1].ToString());

                            object sonuc = m_dbCommand.ExecuteScalar();
                            int sayi = Convert.ToInt32(sonuc);
                            if (sayi > 0)
                                continue;
                        }
                        using (SQLiteCommand m_dbCommand = new SQLiteCommand(m_dbConnection))
                        {
                            m_dbCommand.CommandText = $"INSERT INTO '{tableName}' VALUES ('{dt.Rows[i][0]}','{dt.Rows[i][1]}',{dt.Rows[i][2]},{dt.Rows[i][3]},{dt.Rows[i][4]},{dt.Rows[i][5]})";

                            int affectedRow = m_dbCommand.ExecuteNonQuery();
                            if (affectedRow > 0)
                                affectedRows += affectedRow;
                        }
                    }

                    // SQLite veritabanına kayıt işlemi başarılıysa .cvs dosyası bulunduğu dizinden silinsin
                    FileInfo fi = new FileInfo(e.FullPath);
                    if (affectedRows > 0)
                        fi.Delete();

                    m_dbConnection.Close();
                }
            }
        }

        private static void StartFileSystemWatcher()
        {
            // Grades adında bir klasör oluştur
            if (!Directory.Exists(@"C:\Grades"))
            {
                Directory.CreateDirectory(@"C:\Grades");
            }

            // Grades klasöründeki tüm dosyaları izleyen bir FileSystemWatcher oluştur
            FileSystemWatcher fsw = new FileSystemWatcher(@"C:\Grades\");

            // .csv uzantılı dosyaları filtreler
            fsw.Filter = "*.csv";

            // Dosyaların son erişim ve son yazma zamanlarındaki, ve dosya ya da klasör isimlerindeki değişimi izle
            fsw.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                | NotifyFilters.FileName | NotifyFilters.DirectoryName;

            // Bir dosya oluştuğu zamanki durumu handle et
            fsw.Created += new FileSystemEventHandler(OnCreated);

            // İzlemeye başla
            fsw.EnableRaisingEvents = true;
        }

        private static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            // CSV dosyasını DataTable'a dönüştür
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                DataTable dt = new DataTable();
                foreach (string header in headers)
                {
                    if (header == "ders" || header == "ogrenci_no")
                        dt.Columns.Add(header, typeof(string));
                    else
                        dt.Columns.Add(header, typeof(int));
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                    DataRow dr = dt.NewRow();
                    int currentRowIndex = 0;
                    for (int i = 0; i < headers.Length; i++)
                    {
                        currentRowIndex = i;
                        if (string.IsNullOrEmpty(rows[i])) break;
                        dr[i] = Convert.ChangeType(rows[i], dt.Columns[i].DataType);
                    }
                    if (string.IsNullOrEmpty(rows[currentRowIndex])) break;
                    dt.Rows.Add(dr);
                }
                sr.Close();
                sr.Dispose();
                return dt;
            }
        }
    }
}
