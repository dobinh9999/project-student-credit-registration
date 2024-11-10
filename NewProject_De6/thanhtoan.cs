using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace NewProject_De6
{
    public partial class thanhtoan : Form
    {


        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter;
        DataTable table;


        private Timer expandTimer;
        private int targetWidth = 1370; // Kích thước cuối cùng bạn muốn đạt được
        private int targetHeight = 745; // Chiều cao cuối cùng bạn muốn đạt được
        private int currentWidth;
        private int currentHeight;
        private int step = 5; // Số pixel tăng mỗi lần
        private bool expandingWidth = false; // Biến để xác định xem đang mở rộng theo chiều rộng hay chiều cao
        private bool isExpanding = false;

        public thanhtoan()
        {
            InitializeComponent();

            // Khởi tạo Timer
            expandTimer = new Timer();
            expandTimer.Interval = 10; // 10 milliseconds
            expandTimer.Tick += timer1_Tick;

           

            // Đăng ký sự kiện Click cho Button
            button1.Click += button1_Click;
            button2.Click += button2_Click;


            // Ẩn CheckBox khi form được tạo
            button9.Visible = false;
            button13.Visible = false;
        }
       

   
    private void timer1_Tick(object sender, EventArgs e)
        {
            // Tăng kích thước của form
            if (expandingWidth)
            {
                this.Width += step;
            }
            else
            {

                this.Height += step;

                step = Math.Max(1, step - 1);

            }

            // Kiểm tra xem đã đạt kích thước mong muốn chưa
            if ((expandingWidth && this.Width >= targetWidth) || (!expandingWidth && this.Height >= targetHeight))
            {
                // Dừng Timer khi đã đạt kích thước mong muốn
                expandTimer.Stop();
                isExpanding = false;
            }

        }
        private void ExpandForm(int newWidth, int newHeight, bool expandWidth)
        {
            // Bắt đầu Timer
            expandTimer.Start();
            isExpanding = true;

            // Thiết lập kích thước ban đầu cho mỗi lần nhấn nút
            this.currentWidth = 425; // Đặt lại kích thước ban đầu theo chiều rộng
            this.currentHeight = 316; // Đặt lại kích thước ban đầu theo chiều cao
            this.Width = currentWidth;
            this.Height = currentHeight;

            // Thiết lập targetWidth và targetHeight dựa trên giá trị expandWidth
            targetWidth = expandWidth ? targetWidth : currentWidth;
            targetHeight = expandWidth ? currentHeight : targetHeight;

            // Thiết lập hướng mở rộng (theo chiều rộng hoặc chiều cao)
            expandingWidth = expandWidth;


        }
        private void button1_Click(object sender, EventArgs e)// hiển thị// 
        {

            // Bắt đầu Timer khi nhấn Button1 để mở rộng theo chiều rộng
            if (!isExpanding)
            {
                ExpandForm(this.Width, this.Height, true);
                // Ẩn CheckBox khi form được tạo
                button9.Visible = true;
                button13.Visible = false;
            }



            string searchText = tb_timkiem.Text.Trim();
          

            // Tạo DataTable để lưu trữ dữ liệu tìm kiếm
            DataTable dataTable = new DataTable();

            // Biến để kiểm tra xem có dữ liệu tìm thấy hay không
            bool dataFound = false;

            if (!string.IsNullOrEmpty(searchText))
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(str))
                    {
                        connection.Open();
                        SqlCommand command = connection.CreateCommand();
                        command.CommandText = "SELECT\r\n    ROW_NUMBER() OVER (ORDER BY pdk.MSV) AS 'Số Thứ Tự',\r\n    pdk.MSV AS 'Mã Sinh Viên',\r\n    pdk.Name AS 'Tên Sinh Viên',\r\n    pdk.MaHP AS 'Mã Học Phần',\r\n    MAX(pdk.SoTinChi) AS 'Số Tín Chỉ',\r\n    ISNULL(pdk.SoTinChi * 430000, 0) AS 'Công nợ',\r\n    SUM(ISNULL(pdk.SoTinChi * 430000, 0)) OVER (PARTITION BY pdk.MSV) AS 'Tổng Số Tiền Phải Nộp'\r\nFROM\r\n    qlhp_phieudkHP pdk\r\nLEFT JOIN\r\n    tracuucongno cn ON pdk.MaHP = cn.MaHP\r\nWHERE\r\n    (pdk.MSV LIKE @searchText OR pdk.Name LIKE @searchText)\r\n    AND NOT EXISTS (\r\n        SELECT 1 FROM thanhtoanCN tt WHERE pdk.MSV = tt.MSV\r\n    )\r\nGROUP BY\r\n    pdk.MSV, pdk.Name, pdk.MaHP, pdk.SoTinChi;";

                        command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");
                
                       



                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            DataTable table = new DataTable();
                            dataTable.Columns.Add("Số Thứ Tự");
                            dataTable.Columns.Add("Mã Sinh Viên");
                            dataTable.Columns.Add("Tên Sinh Viên");
                            dataTable.Columns.Add("Mã Học Phần");
                            dataTable.Columns.Add("Số Tín Chỉ");
                            dataTable.Columns.Add("Công nợ");
                            dataTable.Columns.Add("Tổng Số Tiền Phải Nộp");

                            while (reader.Read())
                            {
                                DataRow newRow = dataTable.NewRow();
                                newRow["Số Thứ Tự"] = reader["Số Thứ Tự"].ToString();
                                newRow["Mã Sinh Viên"] = reader["Mã Sinh Viên"].ToString();
                                newRow["Tên Sinh Viên"] = reader["Tên Sinh Viên"].ToString();
                                newRow["Mã Học Phần"] = reader["Mã Học Phần"].ToString();
                                newRow["Số Tín Chỉ"] = reader["Số Tín Chỉ"].ToString();
                                newRow["Công nợ"] = reader["Công nợ"].ToString();
                                newRow["Tổng Số Tiền Phải Nộp"] = reader["Tổng Số Tiền Phải Nộp"].ToString();
                                dataTable.Rows.Add(newRow);
                                dataFound = true;
                                
                            }

                            dataGridView1.DataSource = dataTable;
                            // Đặt biến kiểm tra dữ liệu tìm thấy thành true
                            if (dataFound)
                            {
                                MessageBox.Show("Dữ liệu phù hợp đã được tìm thấy.", "Thông báo!");
                            }
                            // Kiểm tra biến dataFound để thông báo tìm thấy hoặc không tìm thấy
                            if (!dataFound)
                            {
                                MessageBox.Show("Không có dữ liệu phù hợp.", "Thông báo!");
                                button9.Visible = false;
                            }
                        }
                        

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
                }           
                
            }
            else
            {
                MessageBox.Show("Nhập thông tin để tìm kiếm .", "Thông báo!");
                button9.Visible = false;
            }
            
           
        }
        private void button2_Click(object sender, EventArgs e) //  hiển thị 
        {
            if (!isExpanding)
            {
                ExpandForm(this.Width, this.Height, true);
                // Ẩn CheckBox khi form được tạo
                 button9.Visible = false;
                button13.Visible = true;
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();

                    string query = "SELECT\r\n    ROW_NUMBER() OVER (ORDER BY pdk.MSV) AS 'Số Thứ Tự',\r\n    pdk.MSV AS 'Mã Sinh Viên',\r\n    pdk.Name AS 'Tên Sinh Viên',\r\n    pdk.MaHP AS 'Mã Học Phần',\r\n    MAX(pdk.SoTinChi) AS 'Số Tín Chỉ',\r\n    ISNULL(pdk.SoTinChi * 430000, 0) AS 'Công nợ',\r\n    SUM(ISNULL(pdk.SoTinChi * 430000, 0)) OVER (PARTITION BY pdk.MSV) AS 'Tổng Số Tiền Phải Nộp'\r\nFROM\r\n    qlhp_phieudkHP pdk\r\nLEFT JOIN\r\n    tracuucongno cn ON pdk.MaHP = cn.MaHP\r\nWHERE\r\n    NOT EXISTS (\r\n        SELECT 1\r\n        FROM thanhtoanCN tt\r\n        WHERE pdk.MSV = tt.MSV\r\n    )\r\nGROUP BY\r\n    pdk.MSV, pdk.Name, pdk.MaHP, pdk.SoTinChi;";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            using (DataTable table = new DataTable())
                            {
                                adapter.Fill(table);

                                // Gán dữ liệu vào DataGridView1
                                dataGridView1.DataSource = table;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
           

        }
        private void thanhtoan_Load(object sender, EventArgs e)
        {
           
            // TODO: This line of code loads data into the 'project1DataSet2.thanhtoanCN' table. You can move, or remove it, as needed.
            this.thanhtoanCNTableAdapter.Fill(this.project1DataSet2.thanhtoanCN);
            // Thiết lập kích thước ban đầu của form
            this.Width = 425;
            this.Height = 316;

        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void ExpandForm1(int newWidth, int newHeight, bool expandWidth)
        {
            // Bắt đầu Timer
            expandTimer.Start();
            isExpanding = true;

            // Thiết lập kích thước ban đầu cho mỗi lần nhấn nút
            this.currentWidth = 1370; // Đặt lại kích thước ban đầu theo chiều rộng
            this.currentHeight = 528; // Đặt lại kích thước ban đầu theo chiều cao
          
            this.Width = currentWidth;
            this.Height = currentHeight;

            // Thiết lập targetWidth và targetHeight dựa trên giá trị expandWidth
            targetWidth = expandWidth ? targetWidth : currentWidth;
            targetHeight = expandWidth ? currentHeight : targetHeight;

            // Thiết lập hướng mở rộng (theo chiều rộng hoặc chiều cao)
            expandingWidth = expandWidth;


        }
        private void button9_Click(object sender, EventArgs e)// thanh toán 1080 , 510
        {
            if (!isExpanding)
            {
               
                ExpandForm1(this.Width, this.Height, false);
               
            }
        }

        private void ExpandForm2(int newWidth, int newHeight, bool expandWidth)
        {
            // Bắt đầu Timer
            expandTimer.Start();
            isExpanding = true;

            // Thiết lập kích thước ban đầu cho mỗi lần nhấn nút
            this.currentWidth = 1370; // Đặt lại kích thước ban đầu theo chiều rộng
            this.currentHeight = 788; // Đặt lại kích thước ban đầu theo chiều cao

            this.Width = currentWidth;
            this.Height = currentHeight;

            // Thiết lập targetWidth và targetHeight dựa trên giá trị expandWidth
            targetWidth = expandWidth ? targetWidth : currentWidth;
            targetHeight = expandWidth ? currentHeight : targetHeight;

            // Thiết lập hướng mở rộng (theo chiều rộng hoặc chiều cao)
            expandingWidth = expandWidth;


        }
        private void button5_Click(object sender, EventArgs e)// XEM 1080 788
        {
            if (!isExpanding)
            {

                ExpandForm2(this.Width, this.Height, false);

            }
        }


        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            using (SqlConnection connection = new SqlConnection(str))
            {
                try
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();  // Khởi tạo command
                    command.CommandText = "SELECT * FROM thanhtoanCN";

                    SqlDataAdapter adapter = new SqlDataAdapter();  // Khởi tạo adapter
                    adapter.SelectCommand = command;

                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView2.DataSource = table;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
            }
        }
        void loadtext()
        {
            tb_mattcn.ResetText();
            tb_msv.ResetText();
            tb_name.ResetText();
            tb_tongsohp.ResetText();
            tb_tongsotinchi.ResetText();
            tb_sotiendanop.ResetText();

            MessageBox.Show("làm mới thành công.");
        }
        private void button4_Click(object sender, EventArgs e) //luu 1080  
        {

            if (string.IsNullOrWhiteSpace(tb_mattcn.Text) || string.IsNullOrWhiteSpace(tb_msv.Text) || string.IsNullOrWhiteSpace(tb_name.Text) || string.IsNullOrWhiteSpace(tb_tongsohp.Text) || string.IsNullOrWhiteSpace(tb_tongsotinchi.Text) || string.IsNullOrWhiteSpace(tb_sotiendanop.Text))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin.", "thông báo !");
                return;
            }
            using (SqlConnection connection = new SqlConnection(str))
            {
                try
                {
                    connection.Open();
                    command = connection.CreateCommand();

                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM thanhtoanCN WHERE MaTTcn = @MaTT", connection);
                    checkCommand.Parameters.AddWithValue("@MaTT", tb_mattcn.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show($"Mã       {tb_mattcn.Text}đã tồn tại trong cơ sở dữ liệu.\n Vui lòng chọn mã khác.", "thông báo !");
                        return;
                    }

                    command.CommandText = "INSERT INTO thanhtoanCN (MaTTcn, MSV, Name, TongSoHP, TongSoTinChi, SoTienDaNop) VALUES (@MaTT, @MSV, @Name, @TongSoHP, @TongSoTinChi, @SoTienDaNop)";

                    command.Parameters.AddWithValue("@MaTT", tb_mattcn.Text);
                    command.Parameters.AddWithValue("@MSV", tb_msv.Text);
                    command.Parameters.AddWithValue("@Name", tb_name.Text);
                    command.Parameters.AddWithValue("@TongSoHP", Convert.ToInt32(tb_tongsohp.Text));
                    command.Parameters.AddWithValue("@TongSoTinChi", Convert.ToInt32(tb_tongsotinchi.Text));
                    command.Parameters.AddWithValue("@SoTienDaNop", Convert.ToInt32(tb_sotiendanop.Text));

                    command.ExecuteNonQuery();
                    MessageBox.Show($"Đã lưu dữ liệu thanh toán của SINH VIÊN có mã {tb_msv.Text}.");
                    loaddata(); // Cập nhật lại DataGridView sau khi thêm
                  
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private void button6_Click(object sender, EventArgs e)// sua
        {
            using (SqlConnection connection = new SqlConnection(str))
            {

                try
                {
                    connection.Open();
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM thanhtoanCN WHERE MaTTcn = @MaTT", connection);
                    checkCommand.Parameters.AddWithValue("@MaTT", tb_mattcn.Text);
                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        command = connection.CreateCommand();
                        command.CommandText = "UPDATE thanhtoanCN SET MSV = @MSV, Name = @Name, TongSoHP = @TongSoHP, TongSoTinChi = @TongSoTinChi, SoTienDaNop = @SoTienDaNop WHERE MaTTcn = @MaTT";

                        command.Parameters.AddWithValue("@MaTT", tb_mattcn.Text);
                        command.Parameters.AddWithValue("@MSV", tb_msv.Text);
                        command.Parameters.AddWithValue("@Name", tb_name.Text);
                        command.Parameters.AddWithValue("@TongSoHP", Convert.ToInt32(tb_tongsohp.Text));
                        command.Parameters.AddWithValue("@TongSoTinChi", Convert.ToInt32(tb_tongsotinchi.Text));
                        command.Parameters.AddWithValue("@SoTienDaNop", Convert.ToInt32(tb_sotiendanop.Text));

                        command.ExecuteNonQuery();
                        
                        loaddata(); // Cập nhật lại DataGridView sau khi sửa

                        MessageBox.Show("Cập nhật dữ liệu thành công.");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã thanh toán.", "thông báo !");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
             }              
        }
        private void button7_Click(object sender, EventArgs e)// xoa
        {
            using (SqlConnection connection = new SqlConnection(str))
            {
                if (string.IsNullOrWhiteSpace(tb_mattcn.Text) )
                {
                    MessageBox.Show("Vui lòng chọn dữ liệu hoặc nhập thông tin vào mã thanh toán công nợ để xóa.", "thông báo !");
                    return;
                }
                try
                {
                    connection.Open();
                    command = connection.CreateCommand();
                    command.CommandText = "DELETE FROM thanhtoanCN WHERE MaTTcn = @MaTT";

                    command.Parameters.AddWithValue("@MaTT", tb_mattcn.Text);

                    command.ExecuteNonQuery();
                    
                    loaddata(); // Cập nhật lại DataGridView sau khi xóa
                    loadtext(); // Làm mới các TextBox

                    MessageBox.Show("Xóa dữ liệu thành công.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private void button12_Click(object sender, EventArgs e)// tim kiem
        {

            string searchText = tb_timkiem1.Text.Trim();

            // Tạo DataTable để lưu trữ dữ liệu tìm kiếm
            DataTable dataTable = new DataTable();

            // Biến để kiểm tra xem có dữ liệu tìm thấy hay không
            bool dataFound = false;
            using (SqlConnection connection = new SqlConnection(str))
            {
                if (!string.IsNullOrEmpty(searchText))
                {
                    try
                    {
                        connection.Open();
                        command = connection.CreateCommand();
                        command.CommandText = "SELECT * FROM thanhtoanCN WHERE MaTTcn LIKE @searchText OR MSV LIKE @searchText OR Name LIKE @searchText";
                        command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                        SqlDataReader reader = command.ExecuteReader();

                        // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                        dataTable.Columns.Add("MattCN");
                        dataTable.Columns.Add("MSV");
                        dataTable.Columns.Add("Name");
                        dataTable.Columns.Add("Tongsohp");
                        dataTable.Columns.Add("Tongsotinchi");
                        dataTable.Columns.Add("Sotiendanop");

                        // Thêm các cột khác tương tự

                        while (reader.Read())
                        {
                            // Thêm dòng vào DataTable
                            DataRow newRow = dataTable.NewRow();
                            newRow["MattCN"] = reader["MattCN"].ToString();
                            newRow["MSV"] = reader["MSV"].ToString();
                            newRow["Name"] = reader["Name"].ToString();
                            newRow["Tongsohp"] = reader["Tongsohp"].ToString();
                            newRow["Tongsotinchi"] = reader["Tongsotinchi"].ToString();
                            newRow["Sotiendanop"] = reader["Sotiendanop"].ToString();

                            // Thêm các cột khác tương tự
                            dataTable.Rows.Add(newRow);

                            MessageBox.Show(" Đã tìm thấy dữ liệu phù hợp.", "Thông báo!");

                            // Đặt biến kiểm tra dữ liệu tìm thấy thành true
                            dataFound = true;
                        }

                        // Gán DataTable làm nguồn dữ liệu cho DataGridView
                        dataGridView2.DataSource = dataTable;

                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
                    }
                    finally
                    {
                        connection.Close();
                    }

                    // Kiểm tra xem có dữ liệu nào tìm thấy không
                    if (!dataFound)
                    {
                        MessageBox.Show("Không tìm thấy dữ liệu phù hợp.", "Thông báo!");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng nhập thông tin để tìm kiếm.", "Thông báo!");
                }
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }

            tb_mattcn.Text = dataGridView2.Rows[i].Cells[0].Value?.ToString();
            tb_msv.Text = dataGridView2.Rows[i].Cells[1].Value?.ToString();
            tb_name.Text = dataGridView2.Rows[i].Cells[2].Value?.ToString();
            tb_tongsohp.Text = dataGridView2.Rows[i].Cells[3].Value?.ToString();
            tb_tongsotinchi.Text = dataGridView2.Rows[i].Cells[4].Value?.ToString();
            tb_sotiendanop.Text = dataGridView2.Rows[i].Cells[5].Value?.ToString();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            loaddata();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            loadtext();
        }

        private void button10_Click(object sender, EventArgs e)// excel tính công nợ
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_DanhSachCongNoSV.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DANH SÁCH CÔNG NỢ CỦA SINH VIÊN");

                            // Thêm tiêu đề "Danh sách nhân viên"
                            worksheet.Cells["A1"].Value = "DANH SÁCH CÔNG NỢ CỦA SINH VIÊN";
                            worksheet.Cells["A2"].Value = "Thông Tin Chi Tiết";
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 20; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1:G1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:G2"].Merge = true; // Merge các ô từ A1 đến D1                          
                            worksheet.Cells["A3:G3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:G3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:G3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề         
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề
                            worksheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:G1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:G2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:G3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang                          

                            // tạo hàng dọc
                            worksheet.Cells["A1:A3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:G3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

                            // Xuất tiêu đề cột từ DataGridView
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                worksheet.Cells[3, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                            }
                            // Xuất dữ liệu từ DataGridView
                            for (int i = 1; i <= dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 1; j <= dataGridView1.Columns.Count; j++)
                                {
                                    worksheet.Cells[i + 3, j].Value = dataGridView1[j - 1, i - 1].Value;
                                    worksheet.Cells[i + 3, j].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                                    worksheet.Cells[i + 3, j].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                                }
                            }

                            // Tự động điều chỉnh độ rộng của các cột trong Excel sau khi thêm tiêu đề
                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                            package.Save();
                            MessageBox.Show("Xuất dữ liệu thành công!", "Thông báo");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }
        
    }

        private void button8_Click(object sender, EventArgs e)// excel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_DanhSachSVDaTTCONGNO.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DANH SÁCH SINH VIÊN ĐÃ THANH TOÁN ");

                            // Thêm tiêu đề "Danh sách nhân viên"
                            worksheet.Cells["A1"].Value = "DANH SÁCH SINH VIÊN ĐÃ THANH TOÁN CÔNG NỢ";
                            worksheet.Cells["A2"].Value = "Thông Tin Chi Tiết";
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 20; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1:F1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:F2"].Merge = true; // Merge các ô từ A1 đến D1                          
                            worksheet.Cells["A3:F3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:F3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:F3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề         
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề
                            worksheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:F1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:F2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:F3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang                          

                            // tạo hàng dọc
                            worksheet.Cells["A1:A3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:F3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

                            // Xuất tiêu đề cột từ DataGridView
                            for (int i = 0; i < dataGridView2.Columns.Count; i++)
                            {
                                worksheet.Cells[3, i + 1].Value = dataGridView2.Columns[i].HeaderText;
                            }
                            // Xuất dữ liệu từ DataGridView
                            for (int i = 1; i <= dataGridView2.Rows.Count; i++)
                            {
                                for (int j = 1; j <= dataGridView2.Columns.Count; j++)
                                {
                                    worksheet.Cells[i + 3, j].Value = dataGridView2[j - 1, i - 1].Value;
                                    worksheet.Cells[i + 3, j].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                                    worksheet.Cells[i + 3, j].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                                }
                            }

                            // Tự động điều chỉnh độ rộng của các cột trong Excel sau khi thêm tiêu đề
                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                            package.Save();
                            MessageBox.Show("Xuất dữ liệu thành công!", "Thông báo");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }
        }
    }
}
