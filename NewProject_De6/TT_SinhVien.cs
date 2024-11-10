using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using NewProject_De6.FormMain;

namespace NewProject_De6
{

    public partial class TT_SinhVien : Form
    {

        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();

        // kết nối bảng vào form 


        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from qlsv_sinhvien";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }
        public giaodiendangnhap MainFormReference { get; set; }
        private void TT_SinhVien_Load(object sender, EventArgs e)//hiển thị dữ liệu vào dataGriView1
        {

            // TODO: This line of code loads data into the 'project1DataSet.qlsv_sinhvien' table. You can move, or remove it, as needed.
            //this.qlsv_sinhvienTableAdapter.Fill(this.project1DataSet.qlsv_sinhvien);
            connection = new SqlConnection(str); //tạo kết nỗi mới đến chuỗi str
            connection.Open();// mở kết nối khi tải dữ liệu vào form 
            loaddata();// hiện dữ liệu vào form
            connection.Close();//đóng kết nối khi tải dữ liệu xong

            Console.WriteLine("UserRole: " + MainFormReference.UserRole);


            if (MainFormReference != null && MainFormReference.UserRole == "user")
            {
                Console.WriteLine("Disabling buttons for user.");
                // Tắt chức năng thêm và xóa
                button3.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
            }
        }

        public TT_SinhVien()
        {
            InitializeComponent();


        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)// hiển thị dữ dữ liệu vào các textbox khi nhấp vào bảng dataGridView

        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }

            tb_MSV.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_Name.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_GioiTinh.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_NgaySinh.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
            tb_QueQuan.Text = dataGridView1.Rows[i].Cells[4].Value?.ToString();
            tb_TenCV.Text = dataGridView1.Rows[i].Cells[5].Value?.ToString();
            tb_MaNganh.Text = dataGridView1.Rows[i].Cells[6].Value?.ToString();
            tb_MaKhoa.Text = dataGridView1.Rows[i].Cells[7].Value?.ToString();
            tb_MaLop.Text = dataGridView1.Rows[i].Cells[8].Value?.ToString();
        }
        private void button5_Click(object sender, EventArgs e)// chức năng reset
        {
            tb_MSV.ResetText();
            tb_Name.ResetText();
            tb_GioiTinh.ResetText();
            tb_NgaySinh.ResetText();
            tb_QueQuan.ResetText();
            tb_TenCV.ResetText();
            tb_MaNganh.ResetText();
            tb_MaKhoa.ResetText();
            tb_MaLop.ResetText();

            MessageBox.Show("làm mới thành công.", "thông báo !");
        }
        private void button2_Click(object sender, EventArgs e)// chức năng thêm 
        {
            if (string.IsNullOrWhiteSpace(tb_MSV.Text) || string.IsNullOrWhiteSpace(tb_Name.Text) || string.IsNullOrWhiteSpace(tb_GioiTinh.Text) || string.IsNullOrWhiteSpace(tb_NgaySinh.Text) || string.IsNullOrWhiteSpace(tb_QueQuan.Text) || string.IsNullOrWhiteSpace(tb_TenCV.Text) || string.IsNullOrWhiteSpace(tb_MaNganh.Text) || string.IsNullOrWhiteSpace(tb_MaKhoa.Text) || string.IsNullOrWhiteSpace(tb_MaLop.Text))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin.", "thông báo !");
                return;
            }
            try
            {
                connection.Open();

                // Check msv 
                SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_sinhvien WHERE msv = @msv", connection);
                checkCommand.Parameters.AddWithValue("@msv", tb_MSV.Text);
                int count = (int)checkCommand.ExecuteScalar();

                if (count > 0)
                {
                    MessageBox.Show("Mã sinh viên đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã sinh viên khác.", "thông báo !");
                    return;
                }



                command = connection.CreateCommand();
                command.CommandText = "INSERT INTO qlsv_sinhvien VALUES(@msv, @name, @gioiTinh, @ngaySinh, @queQuan, @tenCV, @maNganh, @maKhoa, @maLop)";
                command.Parameters.AddWithValue("@msv", tb_MSV.Text);
                command.Parameters.AddWithValue("@name", tb_Name.Text);
                command.Parameters.AddWithValue("@gioiTinh", tb_GioiTinh.Text);
                command.Parameters.AddWithValue("@ngaySinh", tb_NgaySinh.Value);
                command.Parameters.AddWithValue("@queQuan", tb_QueQuan.Text);
                command.Parameters.AddWithValue("@tenCV", tb_TenCV.Text);
                command.Parameters.AddWithValue("@maNganh", tb_MaNganh.Text);
                command.Parameters.AddWithValue("@maKhoa", tb_MaKhoa.Text);
                command.Parameters.AddWithValue("@maLop", tb_MaLop.Text);
                command.ExecuteNonQuery();
                loaddata();
                MessageBox.Show("Thêm sinh viên thành công.", "thông báo !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "thông báo !");
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
            loaddata();

        }
        private void button4_Click(object sender, EventArgs e)//chức năng sửa
        {
            try
            {
                connection.Open();
                // Kiểm tra xem có mã khoa nào khác với mã khoa hiện tại không nếu có là đang sửa mã khoa 
                SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_sinhvien WHERE msv = @msv", connection);
                checkCommand.Parameters.AddWithValue("@msv", tb_MSV.Text);
                int existingRecords = (int)checkCommand.ExecuteScalar();
                if (existingRecords > 0)
                {
                    command = connection.CreateCommand();
                    command.CommandText = "UPDATE qlsv_sinhvien SET name = @name, gioiTinh = @gioiTinh, ngaySinh = @ngaySinh, queQuan = @queQuan, tenCV = @tenCV, maNganh = @maNganh, maKhoa = @maKhoa, maLop = @maLop WHERE msv = @msv";
                    command.Parameters.AddWithValue("@msv", tb_MSV.Text);
                    command.Parameters.AddWithValue("@name", tb_Name.Text);
                    command.Parameters.AddWithValue("@gioiTinh", tb_GioiTinh.Text);
                    command.Parameters.AddWithValue("@ngaySinh", tb_NgaySinh.Text);
                    command.Parameters.AddWithValue("@queQuan", tb_QueQuan.Text);
                    command.Parameters.AddWithValue("@tenCV", tb_TenCV.Text);
                    command.Parameters.AddWithValue("@maNganh", tb_MaNganh.Text);
                    command.Parameters.AddWithValue("@maKhoa", tb_MaKhoa.Text);
                    command.Parameters.AddWithValue("@maLop", tb_MaLop.Text);
                    command.ExecuteNonQuery();
                    connection.Close();
                    loaddata();
                    MessageBox.Show("Sửa thông tin thành công.");
                }
                else
                {
                    // hiển thị thông báo lỗi
                    MessageBox.Show("Không được sửa mã sinh viên.", "thông báo !");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "thông báo !");
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
            loaddata();
        }
        private void button1_Click(object sender, EventArgs e)//chức năng xóa
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (tb_MSV.Text != "")
            {
                connection.Open();
                command = connection.CreateCommand();
                command.CommandText = "DELETE FROM qlsv_sinhvien WHERE msv = @msv";
                command.Parameters.AddWithValue("@msv", tb_MSV.Text);
                command.ExecuteNonQuery();
                connection.Close();
                loaddata();

                MessageBox.Show("xóa sinh viên thành công.", "thông báo !");

                tb_MSV.ResetText();
                tb_Name.ResetText();
                tb_GioiTinh.ResetText();
                tb_NgaySinh.ResetText();
                tb_QueQuan.ResetText();
                tb_TenCV.ResetText();
                tb_MaNganh.ResetText();
                tb_MaKhoa.ResetText();
                tb_MaLop.ResetText();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn sinh viên cần xóa.", "thông báo !");

            }

            connection.Close();
        }
        private void button3_Click(object sender, EventArgs e) // chức năng tìm kiếm 
        {
            string searchText = tb_timkiem.Text.Trim();

            // Tạo DataTable để lưu trữ dữ liệu tìm kiếm
            DataTable dataTable = new DataTable();

            // Biến để kiểm tra xem có dữ liệu tìm thấy hay không
            bool dataFound = false;

            if (!string.IsNullOrEmpty(searchText))
            {
                try
                {
                    connection.Open();
                    command = connection.CreateCommand();
                    command.CommandText = "SELECT * FROM qlsv_sinhvien WHERE msv LIKE @searchText OR Name LIKE @searchText OR gioitinh LIKE @searchText  OR queQuan LIKE @searchText  OR NgaySinh LIKE @searchText OR CONVERT(VARCHAR, ngaySinh, 103) LIKE @searchText OR maKhoa LIKE @searchText OR maLop LIKE @searchText OR maNganh LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("msv");
                    dataTable.Columns.Add("name");
                    dataTable.Columns.Add("gioitinh");
                    dataTable.Columns.Add("ngaySinh");
                    dataTable.Columns.Add("queQuan");
                    dataTable.Columns.Add("tenCV");
                    dataTable.Columns.Add("maNganh");
                    dataTable.Columns.Add("maKhoa");
                    dataTable.Columns.Add("maLop");
                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["msv"] = reader["msv"].ToString();
                        newRow["name"] = reader["name"].ToString();
                        newRow["gioitinh"] = reader["gioitinh"].ToString();
                        newRow["ngaySinh"] = reader["ngaySinh"].ToString();
                        newRow["queQuan"] = reader["queQuan"].ToString();
                        newRow["tenCV"] = reader["tenCV"].ToString();
                        newRow["maNganh"] = reader["maNganh"].ToString();
                        newRow["maKhoa"] = reader["maKhoa"].ToString();
                        newRow["maLop"] = reader["maLop"].ToString();

                        // Thêm các cột khác tương tự
                        dataTable.Rows.Add(newRow);

                        // Đặt biến kiểm tra dữ liệu tìm thấy thành true
                        dataFound = true;
                    }



                    // Gán DataTable làm nguồn dữ liệu cho DataGridView
                    dataGridView1.DataSource = dataTable;

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
                MessageBox.Show("Vui lòng nhập yêu cầu để tìm kiếm thông tin.", "Thông báo!");
            }
        }
        private void button7_Click(object sender, EventArgs e)// chức năng hiển thị lại sau khi tìm kiếm xong 
        {
            loaddata();
        }
        private void button6_Click(object sender, EventArgs e)// chức năng in file excel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_sinhvien.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_SINH VIÊN");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh sách sinh viên";
                            worksheet.Cells["A1:I1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:I2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A3:I3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:I3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:I3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:I1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:I2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:I3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A109:I109"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A109"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:I109"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

                            // Xuất tiêu đề cột
                            for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                            {
                                worksheet.Cells[3, i].Value = dataGridView1.Columns[i - 1].HeaderText;
                                worksheet.Cells[3, i].Style.Font.Bold = true; // Đặt đậm cho tiêu đề cột
                                worksheet.Cells[3, i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề cột
                            }

                            // Xuất dữ liệu từ DataGridView
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                                {
                                    worksheet.Cells[i + 4, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                                }
                            }

                            // Tự động điều chỉnh độ rộng của các cột trong Excel
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

        private void qlsvsinhvienBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }


        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }





        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tb_GioiTinh_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)// nhập dữ liệu từ excel
        {
            MessageBox.Show("Bạn muốn xem dữ liệu dữ liệu trong excel.", "Thông báo!");
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;

                        // Load dữ liệu từ Excel
                        DataTable excelData = LoadDataFromExcel(filePath);

                        // Kiểm tra nếu có dữ liệu
                        if (excelData != null && excelData.Rows.Count > 0)
                        {
                            // Gán dữ liệu từ DataTable của Excel vào DataTable của DataGridView
                            table.Clear();
                            foreach (DataRow row in excelData.Rows)
                            {
                                table.Rows.Add(row.ItemArray);
                            }

                            // Hiển thị dữ liệu trên DataGridView
                            dataGridView1.DataSource = table;

                            

                            MessageBox.Show("Đã nhập dữ liệu từ Excel vào DataGridView.", "Thông báo");
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy dữ liệu trong tệp Excel.", "Thông báo");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }
        }
        private DataTable LoadDataFromExcel(string filePath)
        {
            DataTable excelData = new DataTable();

            try
            {
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    // Tạo cấu trúc DataTable dựa trên số cột trong Excel
                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        excelData.Columns.Add($"Column{i}", typeof(string));
                    }

                    // Đọc dữ liệu từ Excel và thêm vào DataTable
                    for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                    {
                        DataRow dataRow = excelData.NewRow();
                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString();
                        }
                        excelData.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }

            return excelData;
        }
        private void SaveDataToDatabase(string filePath)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 1; row <= rowCount; row++) // Bắt đầu từ dòng thứ 2 để bỏ qua dòng tiêu đề
                        {
                            // Lấy giá trị từ cột trong Excel
                            string msv = worksheet.Cells[row, 1].Value?.ToString();
                            string name = worksheet.Cells[row, 2].Value?.ToString();
                            string gioiTinh = worksheet.Cells[row, 3].Value?.ToString();
                            string ngaySinhString = worksheet.Cells[row, 4].Value?.ToString();
                            string queQuan = worksheet.Cells[row, 5].Value?.ToString();
                            string tenCV = worksheet.Cells[row, 6].Value?.ToString();
                            string maNganh = worksheet.Cells[row, 7].Value?.ToString();
                            string maKhoa = worksheet.Cells[row, 8].Value?.ToString();
                            string maLop = worksheet.Cells[row, 9].Value?.ToString();

                            // Kiểm tra xem có phải là dòng đầu tiên hay không
                            if (!string.IsNullOrEmpty(msv))
                            {
                                // Chuyển đổi giá trị ngày sinh từ kiểu ký tự sang kiểu DateTime
                                DateTime ngaySinh;
                                if (DateTime.TryParse(ngaySinhString, out ngaySinh))
                                {
                                    // Tạo câu lệnh SQL INSERT
                                    string sqlCommandText = "INSERT INTO qlsv_sinhvien (msv, name, gioiTinh, ngaySinh, queQuan, tenCV, maNganh, maKhoa, maLop) VALUES (@msv, @name, @gioiTinh, @ngaySinh, @queQuan, @tenCV, @maNganh, @maKhoa, @maLop)";

                                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                                    {
                                        // Gán giá trị từ Excel vào các tham số của câu lệnh SQL
                                        command.Parameters.AddWithValue("@msv", msv);
                                        command.Parameters.AddWithValue("@name", name);
                                        command.Parameters.AddWithValue("@gioiTinh", gioiTinh);
                                        command.Parameters.AddWithValue("@ngaySinh", ngaySinh);
                                        command.Parameters.AddWithValue("@queQuan", queQuan);
                                        command.Parameters.AddWithValue("@tenCV", tenCV);
                                        command.Parameters.AddWithValue("@maNganh", maNganh);
                                        command.Parameters.AddWithValue("@maKhoa", maKhoa);
                                        command.Parameters.AddWithValue("@maLop", maLop);

                                        command.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    // Xử lý trường hợp không chuyển đổi được giá trị ngày sinh
                                    // (Ví dụ: thông báo lỗi hoặc thực hiện xử lý khác)
                                }
                            }
                        }
                    }

                    MessageBox.Show("Dữ liệu đã được lưu vào cơ sở dữ liệu thành công.", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu vào cơ sở dữ liệu: " + ex.Message, "Thông báo lỗi");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MessageBox.Show("chọn tệp excel , ấn open để lấy dữ liệu cần lưu.", "Thông báo!");
            try
            {
                // Sử dụng OpenFileDialog để chọn tệp Excel
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Gọi phương thức SaveDataToDatabase để lưu dữ liệu vào cơ sở dữ liệu từ tệp Excel được chọn
                        SaveDataToDatabase(openFileDialog.FileName);
                    }
                    loaddata();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }
        }
    }
}
