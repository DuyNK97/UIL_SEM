﻿using QRCoder;
using System;
using System.Data;
using System.Data.SQLite;
using System.Configuration;
using CIM.Class;
using System.IO;

namespace CIM
{
    public class SqlLite
    {
        private static SqlLite _instance;
        private static readonly object _lock = new object();
        private SQLiteConnection _connection;
        // string ConnectionString = "Data Source=C:\\Users\\SEM_PROJECT_CIMPC\\Desktop\\Program\\UILPJ_usingSEMIDB\\CIM\\CIM\\Database\\SEMPJ.db";
        string ConnectionString = "Data Source=C:\\APP\\CIMDB\\SEM.db";
        //private const string ConnectionString = ConfigurationManager.AppSettings["Connection"].ToString();

        private SqlLite()
        {
            InitializeConnection();

        }

        public void InitializeConnection()
        {
            try
            {
                string dbFilePath = ConnectionString.Split('=')[1].Trim();

                string directoryPath = Path.GetDirectoryName(dbFilePath);

                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(dbFilePath));
                }

                if (!File.Exists(dbFilePath))
                {
                    SQLiteConnection.CreateFile(dbFilePath);
                }

                _connection = new SQLiteConnection(ConnectionString);
                _connection.Open();
                CreateTables(); // Tạo các bảng khi khởi tạo kết nối
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error can not create file database, error: {ex.Message}");
            }
        }

        public static SqlLite Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new SqlLite();
                    }
                    return _instance;
                }
            }
        }

        // chi tao bang mot lan
        private void CreateTables()
        {
            CreateSEM_DATATable();
            CreateBOX1_DATATable();
            CreateBOX2_DATATable();
            CreateBOX3_DATATable();
            CreateBOX4_DATATable();
        }

        // Các phương thức tạo bảng vẫn giữ nguyên như trước
        private void CreateSEM_DATATable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS SEM_DATA (
                    TOPHOUSING TEXT  ,
                    BOX1_GLUE_AMOUNT TEXT,
                    BOX1_GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    INSULATOR_BAR_CODE TEXT,
                    BOX1_GLUE_OVERFLOW_VISION TEXT,
                    BOX2_GLUE_AMOUNT TEXT,
                    BOX2_GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    FPCB_BAR_CODE TEXT,
                    BOX2_GLUE_OVERFLOW_VISION TEXT,
                    BOX1_HEATED_AIR_CURING TEXT,
                    BOX2_HEATED_AIR_CURING TEXT,
                    BOX3_DISTANCE TEXT,
                    BOX3_GLUE_AMOUNT TEXT,
                    BOX3_GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    BOX3_GLUE_OVERFLOW_VISION TEXT,
                    BOX4_AIR_LEAKAGE_TEST_DETAIL TEXT,
                    BOX3_HEATED_AIR_CURING TEXT,
                    BOX4_TIGHTNESS_AND_LOCATION_VISION TEXT,
                    BOX4_HEIGHT_PARALLELISM TEXT,
                    BOX4_RESISTANCE TEXT,
                    BOX4_AIR_LEAKAGE_TEST_RESULT TEXT,
                    BOX4_TestTime TEXT,
                    Remark TEXT
                );";

            ExecuteNonQuery(createTableQuery);
        }

        private void CreateBOX1_DATATable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS BOX1_DATA (
                    TOPHOUSING TEXT,
                    GLUE_AMOUNT TEXT,
                    GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    INSULATOR_BAR_CODE TEXT,
                    GLUE_OVERFLOW_VISION TEXT,
                    TestTime DATETIME,
                    PRIMARY KEY (TOPHOUSING, TestTime)
                );";

            ExecuteNonQuery(createTableQuery);
        }

        private void CreateBOX2_DATATable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS BOX2_DATA (
                    TOPHOUSING TEXT,
                    GLUE_AMOUNT TEXT,
                    GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    FPCB_BAR_CODE TEXT,
                    GLUE_OVERFLOW_VISION TEXT,
                    BOX1_HEATED_AIR_CURING TEXT,
                    TestTime DATETIME,
                    PRIMARY KEY (TOPHOUSING, TestTime)
                );";

            ExecuteNonQuery(createTableQuery);
        }

        private void CreateBOX3_DATATable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS BOX3_DATA (
                    TOPHOUSING TEXT,
                    BOX2_HEATED_AIR_CURING TEXT,
                    DISTANCE TEXT,
                    GLUE_AMOUNT TEXT,
                    GLUE_DISCHARGE_VOLUME_VISION TEXT,
                    BOX3_GLUE_OVERFLOW_VISION TEXT,
                    TestTime DATETIME,
                    PRIMARY KEY (TOPHOUSING, TestTime)
                );";

            ExecuteNonQuery(createTableQuery);
        }

        private void CreateBOX4_DATATable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS BOX4_DATA (
                    TOPHOUSING TEXT,
                    AIR_LEAKAGE_TEST_DETAIL TEXT,
                    BOX3_HEATED_AIR_CURING TEXT,
                    TIGHTNESS_AND_LOCATION_VISION TEXT,
                    HEIGHT_PARALLELISM TEXT,
                    RESISTANCE TEXT,
                    AIR_LEAKAGE_TEST_RESULT TEXT,
                    TestTime DATETIME,
                    PRIMARY KEY (TOPHOUSING, TestTime)
                );";

            ExecuteNonQuery(createTableQuery);
        }
        private void ExecuteNonQuery(string query)
        {
            using (var command = new SQLiteCommand(query, _connection))
            {
                command.ExecuteNonQuery();
            }
        }

        public bool CheckQRcode(string QRcode)
        {
            DataSet ds = new DataSet();

            try
            {
                string sql = "SELECT * FROM SEM_DATA WHERE TOPHOUSING=@qrCode"; // Sửa thành TOPHOUSING=@qrCode
                using (SQLiteConnection conn = new SQLiteConnection(ConnectionString))
                {
                    conn.Open();

                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@qrCode", QRcode);

                        using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd))
                        {
                            adapter.Fill(ds);
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            if (chkDS(ds))
            {
                return false;
            }
            return true;
        }

        public DataSet SearchData(string QRcode )
        {
            DataSet ds = new DataSet();

            try
            {
                string sql = "SELECT * FROM SEM_DATA";
                if ( !string.IsNullOrEmpty( QRcode) )
                {
                    sql += " WHERE TOPHOUSING LIKE @qrCode"; // Sử dụng LIKE để tìm kiếm mẫu chuỗi
                }

                using (SQLiteConnection conn = new SQLiteConnection(ConnectionString))
                {
                    conn.Open();

                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        if (!string.IsNullOrEmpty(QRcode))
                        {
                            cmd.Parameters.AddWithValue("@qrCode", "%" + QRcode + "%"); // Thêm dấu % cho LIKE
                        }

                        using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd))
                        {
                            adapter.Fill(ds);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return ds;
        }

        private bool chkDS(DataSet ds)
        {
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        return true;
                    }
                }
            }
            return false;
        }


        public void InsertSEM_DATA(EXCELDATA data)
        {
            string insertQuery = @"
                INSERT INTO SEM_DATA (
                    TOPHOUSING, 
                    BOX1_GLUE_AMOUNT, 
                    BOX1_GLUE_DISCHARGE_VOLUME_VISION, 
                    INSULATOR_BAR_CODE, 
                    BOX1_GLUE_OVERFLOW_VISION, 
                    BOX2_GLUE_AMOUNT, 
                    BOX2_GLUE_DISCHARGE_VOLUME_VISION, 
                    FPCB_BAR_CODE, 
                    BOX2_GLUE_OVERFLOW_VISION, 
                    BOX1_HEATED_AIR_CURING, 
                    BOX2_HEATED_AIR_CURING, 
                    BOX3_DISTANCE, 
                    BOX3_GLUE_AMOUNT, 
                    BOX3_GLUE_DISCHARGE_VOLUME_VISION, 
                    BOX3_GLUE_OVERFLOW_VISION, 
                    BOX4_AIR_LEAKAGE_TEST_DETAIL, 
                    BOX3_HEATED_AIR_CURING, 
                    BOX4_TIGHTNESS_AND_LOCATION_VISION, 
                    BOX4_HEIGHT_PARALLELISM, 
                    BOX4_RESISTANCE, 
                    BOX4_AIR_LEAKAGE_TEST_RESULT, 
                    BOX4_TestTime
                )
                VALUES (
                    @TOPHOUSING, 
                    @BOX1_GLUE_AMOUNT, 
                    @BOX1_GLUE_DISCHARGE_VOLUME_VISION, 
                    @INSULATOR_BAR_CODE, 
                    @BOX1_GLUE_OVERFLOW_VISION, 
                    @BOX2_GLUE_AMOUNT, 
                    @BOX2_GLUE_DISCHARGE_VOLUME_VISION, 
                    @FPCB_BAR_CODE, 
                    @BOX2_GLUE_OVERFLOW_VISION, 
                    @BOX1_HEATED_AIR_CURING, 
                    @BOX2_HEATED_AIR_CURING, 
                    @BOX3_DISTANCE, 
                    @BOX3_GLUE_AMOUNT, 
                    @BOX3_GLUE_DISCHARGE_VOLUME_VISION, 
                    @BOX3_GLUE_OVERFLOW_VISION, 
                    @BOX4_AIR_LEAKAGE_TEST_DETAIL, 
                    @BOX3_HEATED_AIR_CURING, 
                    @BOX4_TIGHTNESS_AND_LOCATION_VISION, 
                    @BOX4_HEIGHT_PARALLELISM, 
                    @BOX4_RESISTANCE, 
                    @BOX4_AIR_LEAKAGE_TEST_RESULT, 
                    @BOX4_TestTime
                )";

            using (var command = new SQLiteCommand(insertQuery, _connection))
            {
                command.Parameters.AddWithValue("@TOPHOUSING", data.TOPHOUSING);
                command.Parameters.AddWithValue("@BOX1_GLUE_AMOUNT", data.BOX1_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX1_GLUE_DISCHARGE_VOLUME_VISION", data.BOX1_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@INSULATOR_BAR_CODE", data.INSULATOR_BAR_CODE);
                command.Parameters.AddWithValue("@BOX1_GLUE_OVERFLOW_VISION", data.BOX1_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX2_GLUE_AMOUNT", data.BOX2_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX2_GLUE_DISCHARGE_VOLUME_VISION", data.BOX2_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@FPCB_BAR_CODE", data.FPCB_BAR_CODE);
                command.Parameters.AddWithValue("@BOX2_GLUE_OVERFLOW_VISION", data.BOX2_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX1_HEATED_AIR_CURING", data.BOX1_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX2_HEATED_AIR_CURING", data.BOX2_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX3_DISTANCE", data.BOX3_DISTANCE);
                command.Parameters.AddWithValue("@BOX3_GLUE_AMOUNT", data.BOX3_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX3_GLUE_DISCHARGE_VOLUME_VISION", data.BOX3_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@BOX3_GLUE_OVERFLOW_VISION", data.BOX3_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX4_AIR_LEAKAGE_TEST_DETAIL", data.BOX4_AIR_LEAKAGE_TEST_DETAIL);
                command.Parameters.AddWithValue("@BOX3_HEATED_AIR_CURING", data.BOX3_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX4_TIGHTNESS_AND_LOCATION_VISION", data.BOX4_TIGHTNESS_AND_LOCATION_VISION);
                command.Parameters.AddWithValue("@BOX4_HEIGHT_PARALLELISM", data.BOX4_HEIGHT_PARALLELISM);
                command.Parameters.AddWithValue("@BOX4_RESISTANCE", data.BOX4_RESISTANCE);
                command.Parameters.AddWithValue("@BOX4_AIR_LEAKAGE_TEST_RESULT", data.BOX4_AIR_LEAKAGE_TEST_RESULT);
                command.Parameters.AddWithValue("@BOX4_TestTime", DateTime.Now);

                command.ExecuteNonQuery();
            }
        }
        public void InsertSEM_DATA(EXCELDATA data,string  remark)
        {
            string insertQuery = @"
                INSERT INTO SEM_DATA (
                    TOPHOUSING, 
                    BOX1_GLUE_AMOUNT, 
                    BOX1_GLUE_DISCHARGE_VOLUME_VISION, 
                    INSULATOR_BAR_CODE, 
                    BOX1_GLUE_OVERFLOW_VISION, 
                    BOX2_GLUE_AMOUNT, 
                    BOX2_GLUE_DISCHARGE_VOLUME_VISION, 
                    FPCB_BAR_CODE, 
                    BOX2_GLUE_OVERFLOW_VISION, 
                    BOX1_HEATED_AIR_CURING, 
                    BOX2_HEATED_AIR_CURING, 
                    BOX3_DISTANCE, 
                    BOX3_GLUE_AMOUNT, 
                    BOX3_GLUE_DISCHARGE_VOLUME_VISION, 
                    BOX3_GLUE_OVERFLOW_VISION, 
                    BOX4_AIR_LEAKAGE_TEST_DETAIL, 
                    BOX3_HEATED_AIR_CURING, 
                    BOX4_TIGHTNESS_AND_LOCATION_VISION, 
                    BOX4_HEIGHT_PARALLELISM, 
                    BOX4_RESISTANCE, 
                    BOX4_AIR_LEAKAGE_TEST_RESULT, 
                    BOX4_TestTime,Remark
                )
                VALUES (
                    @TOPHOUSING, 
                    @BOX1_GLUE_AMOUNT, 
                    @BOX1_GLUE_DISCHARGE_VOLUME_VISION, 
                    @INSULATOR_BAR_CODE, 
                    @BOX1_GLUE_OVERFLOW_VISION, 
                    @BOX2_GLUE_AMOUNT, 
                    @BOX2_GLUE_DISCHARGE_VOLUME_VISION, 
                    @FPCB_BAR_CODE, 
                    @BOX2_GLUE_OVERFLOW_VISION, 
                    @BOX1_HEATED_AIR_CURING, 
                    @BOX2_HEATED_AIR_CURING, 
                    @BOX3_DISTANCE, 
                    @BOX3_GLUE_AMOUNT, 
                    @BOX3_GLUE_DISCHARGE_VOLUME_VISION, 
                    @BOX3_GLUE_OVERFLOW_VISION, 
                    @BOX4_AIR_LEAKAGE_TEST_DETAIL, 
                    @BOX3_HEATED_AIR_CURING, 
                    @BOX4_TIGHTNESS_AND_LOCATION_VISION, 
                    @BOX4_HEIGHT_PARALLELISM, 
                    @BOX4_RESISTANCE, 
                    @BOX4_AIR_LEAKAGE_TEST_RESULT, 
                    @BOX4_TestTime,
                    @Remark

                )";

            using (var command = new SQLiteCommand(insertQuery, _connection))
            {
                command.Parameters.AddWithValue("@TOPHOUSING", data.TOPHOUSING);
                command.Parameters.AddWithValue("@BOX1_GLUE_AMOUNT", data.BOX1_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX1_GLUE_DISCHARGE_VOLUME_VISION", data.BOX1_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@INSULATOR_BAR_CODE", data.INSULATOR_BAR_CODE);
                command.Parameters.AddWithValue("@BOX1_GLUE_OVERFLOW_VISION", data.BOX1_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX2_GLUE_AMOUNT", data.BOX2_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX2_GLUE_DISCHARGE_VOLUME_VISION", data.BOX2_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@FPCB_BAR_CODE", data.FPCB_BAR_CODE);
                command.Parameters.AddWithValue("@BOX2_GLUE_OVERFLOW_VISION", data.BOX2_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX1_HEATED_AIR_CURING", data.BOX1_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX2_HEATED_AIR_CURING", data.BOX2_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX3_DISTANCE", data.BOX3_DISTANCE);
                command.Parameters.AddWithValue("@BOX3_GLUE_AMOUNT", data.BOX3_GLUE_AMOUNT);
                command.Parameters.AddWithValue("@BOX3_GLUE_DISCHARGE_VOLUME_VISION", data.BOX3_GLUE_DISCHARGE_VOLUME_VISION);
                command.Parameters.AddWithValue("@BOX3_GLUE_OVERFLOW_VISION", data.BOX3_GLUE_OVERFLOW_VISION);
                command.Parameters.AddWithValue("@BOX4_AIR_LEAKAGE_TEST_DETAIL", data.BOX4_AIR_LEAKAGE_TEST_DETAIL);
                command.Parameters.AddWithValue("@BOX3_HEATED_AIR_CURING", data.BOX3_HEATED_AIR_CURING);
                command.Parameters.AddWithValue("@BOX4_TIGHTNESS_AND_LOCATION_VISION", data.BOX4_TIGHTNESS_AND_LOCATION_VISION);
                command.Parameters.AddWithValue("@BOX4_HEIGHT_PARALLELISM", data.BOX4_HEIGHT_PARALLELISM);
                command.Parameters.AddWithValue("@BOX4_RESISTANCE", data.BOX4_RESISTANCE);
                command.Parameters.AddWithValue("@BOX4_AIR_LEAKAGE_TEST_RESULT", data.BOX4_AIR_LEAKAGE_TEST_RESULT);
                command.Parameters.AddWithValue("@BOX4_TestTime", DateTime.Now);
                command.Parameters.AddWithValue("@Remark", remark);
                command.ExecuteNonQuery();
            }
        }
        public bool UpdateTrayQRcode(string topHousing, string trayCode)
        {
            try
            {
                string sql = "UPDATE SEM_DATA SET qr_code = @trayCode WHERE TOPHOUSING = @topHousing";

                using (SQLiteConnection conn = new SQLiteConnection(ConnectionString))
                {
                    conn.Open();

                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@topHousing", topHousing);
                        cmd.Parameters.AddWithValue("@trayCode", trayCode);

                        int rowsUpdated = cmd.ExecuteNonQuery();

                        // Check if any rows were updated
                        if (rowsUpdated == 0)
                        {
                            Console.WriteLine("No rows updated.");
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return false;
            }
        }

    }
}
