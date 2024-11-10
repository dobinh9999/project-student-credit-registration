using NewProject_De6.FormMain;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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

namespace NewProject_De6
{
    public partial class PhieuDK : Form
    {
        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();


        public PhieuDK()
        {
            InitializeComponent();
        }
        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from qlhp_phieudkHP";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }
        public giaodiendangnhap MainFormReference { get; set; }
        private void PhieuDK_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet.qlhp_phieudkHP' table. You can move, or remove it, as needed.
            //this.qlhp_phieudkHPTableAdapter.Fill(this.project1DataSet.qlhp_phieudkHP);
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
                button2.Enabled = false;
                button6.Enabled = false;
                button8.Enabled = false;
            }
        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }

            tb_ma.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_msv.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_ten.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_ngaydk.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
            tb_hocky.Text = dataGridView1.Rows[i].Cells[4].Value?.ToString();
            tb_mahp.Text = dataGridView1.Rows[i].Cells[5].Value?.ToString();
            tb_tenhp.Text = dataGridView1.Rows[i].Cells[6].Value?.ToString();
            tb_tinchi.Text = dataGridView1.Rows[i].Cells[7].Value?.ToString();

        }

        private void button1_Click(object sender, EventArgs e)// CHỨC NĂNG THÊM 
        {
            if (string.IsNullOrWhiteSpace(tb_ma.Text) || string.IsNullOrWhiteSpace(tb_msv.Text) || string.IsNullOrWhiteSpace(tb_ten.Text) || string.IsNullOrWhiteSpace(tb_ngaydk.Text) || string.IsNullOrWhiteSpace(tb_hocky.Text) || string.IsNullOrWhiteSpace(tb_mahp.Text) || string.IsNullOrWhiteSpace(tb_tenhp.Text) || string.IsNullOrWhiteSpace(tb_tinchi.Text))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin.", "thông báo !");
                return;
            }

            using (SqlConnection connection = new SqlConnection(str))
            {
                try
                {

                    connection.Open();

                    // Check manganh
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlhp_phieudkHP WHERE MaPDK = @MaPDK", connection);
                    checkCommand.Parameters.AddWithValue("@MaPDK", tb_ma.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Mã phiếu đăng ký đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã phiếu khác.", "thông báo !");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO qlhp_phieudkHP (MaPDK, msv, Name, ngaydk, hocky, mahp, TenHP, SoTinChi) VALUES (@MaPDK, @msv, @Name, @ngaydk, @hocky, @mahp, @TenHP, @SoTinChi)", connection);
                    command.Parameters.AddWithValue("@MaPDK", tb_ma.Text);
                    command.Parameters.AddWithValue("@msv", tb_msv.Text);
                    command.Parameters.AddWithValue("@Name", tb_ten.Text);
                    command.Parameters.AddWithValue("@ngaydk", tb_ngaydk.Text);
                    command.Parameters.AddWithValue("@hocky", tb_hocky.Text);
                    command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
                    command.Parameters.AddWithValue("@SoTinChi", tb_tinchi.Text);
                    int rowsAffected = command.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Thêm thành công.", "thông báo !");
                        loaddata(); // Gọi hàm loaddata() để cập nhật DataGridView
                    }
                    else
                    {
                        MessageBox.Show("Thêm không thành công. Vui lòng kiểm tra lại dữ liệu.", "thông báo !");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "thông báo !");
                }
                finally
                {
                    connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
                }

            }
        }
        private void button3_Click(object sender, EventArgs e)// CHỨC NĂNG XÓA 
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(str))
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand("DELETE FROM qlhp_phieudkHP WHERE MaPDK = @ma", connection);
                        command.Parameters.AddWithValue("@ma", tb_ma.Text);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Xóa dữ liệu thành công.", "thông báo !");
                        tb_ma.ResetText();
                        tb_msv.ResetText();
                        tb_ten.ResetText();
                        tb_ngaydk.ResetText();
                        tb_hocky.ResetText();
                        tb_mahp.ResetText();
                        tb_tenhp.ResetText();
                        tb_tinchi.ResetText();
                    }
                    loaddata(); // Reload data after deleting
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "thông báo !");
                }
                finally
                {
                    connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)//chức năng sửa 
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();
                    // Kiểm tra xem có mã  nào khác với mã khoa hiện tại không nếu có là đang sửa mã  
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlhp_phieudkHP WHERE MaPDK = @MaPDK", connection);
                    checkCommand.Parameters.AddWithValue("@MaPDK", tb_ma.Text);
                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE qlhp_phieudkHP SET msv = @msv, Name = @Name, ngaydk =@ngaydk, hocky=@hocky,mahp=@mahp,TenHP=@TenHP,sotinchi=@sotinchi  WHERE MaPDK = @MaPDK", connection);
                        command.Parameters.AddWithValue("@MaPDK", tb_ma.Text);
                        command.Parameters.AddWithValue("@msv", tb_msv.Text);
                        command.Parameters.AddWithValue("@Name", tb_ten.Text);
                        command.Parameters.AddWithValue("@ngaydk", tb_ngaydk.Text);
                        command.Parameters.AddWithValue("@hocky", tb_hocky.Text);
                        command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                        command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
                        command.Parameters.AddWithValue("@sotinchi", tb_tinchi.Text);
                        command.ExecuteNonQuery();
                        connection.Close();
                        loaddata();

                        MessageBox.Show("Sửa thông tin thành công.", "thông báo !");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã phiếu.", "thông báo !");
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "thông báo !");
            }
            finally
            {
                connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
            }
        }
        private void button4_Click(object sender, EventArgs e)// chức năng làm mới 
        {
            tb_ma.ResetText();
            tb_msv.ResetText();
            tb_ten.ResetText();
            tb_ngaydk.ResetText();
            tb_hocky.ResetText();
            tb_mahp.ResetText();
            tb_tenhp.ResetText();
            tb_tinchi.ResetText();
            MessageBox.Show("làm mới thành công.", "thông báo !");
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)// hiển thị lại dữ liệu 
        {
            loaddata();
        }

        private void button5_Click(object sender, EventArgs e)// chức năng tìm kiếm 
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
                    command.CommandText = "SELECT * FROM qlhp_phieudkHP WHERE MaPDK LIKE @searchText OR msv LIKE @searchText OR Name LIKE @searchText  OR ngaydk LIKE @searchText  OR ngaydk LIKE @searchText OR CONVERT(VARCHAR, ngaydk, 103) LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("MaPDK");
                    dataTable.Columns.Add("msv");
                    dataTable.Columns.Add("Name");
                    dataTable.Columns.Add("ngaydk");
                    dataTable.Columns.Add("hocky");
                    dataTable.Columns.Add("mahp");
                    dataTable.Columns.Add("TenHP");
                    dataTable.Columns.Add("sotinchi");

                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["MaPDK"] = reader["MaPDK"].ToString();
                        newRow["msv"] = reader["msv"].ToString();
                        newRow["Name"] = reader["Name"].ToString();
                        newRow["ngaydk"] = reader["ngaydk"].ToString();
                        newRow["hocky"] = reader["hocky"].ToString();
                        newRow["mahp"] = reader["mahp"].ToString();
                        newRow["TenHP"] = reader["TenHP"].ToString();
                        newRow["sotinchi"] = reader["sotinchi"].ToString();


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

        private void button6_Click(object sender, EventArgs e)// chức năng xuất excel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_PHIEUDK.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_PHIEUDK");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Phiếu đăng ký";
                            worksheet.Cells["A1:H1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:H2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:H3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:C3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:H1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:H2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:H3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A109:H109"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A109"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:H109"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e) // chức năng xuất excel báo cáo
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_BaoCao.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Báo cáo");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "BÁO CÁO THỐNG KÊ";
                            worksheet.Cells["A2"].Value = "Sinh Viên Đã Đăng Ký Thành Công";
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 20; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1:H1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:H2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A3:H3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:H3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:H3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề         
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề
                            worksheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:H1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:H2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:H3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang                          

                            // tạo hàng dọc
                            worksheet.Cells["A1:A3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:H3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

                            int startRow = 3;
                            int startColumn = 1;

                            // Xuất tiêu đề cột và tạo đường viền cho hàng ngang
                            for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                            {
                                worksheet.Cells[startRow, startColumn].Value = dataGridView1.Columns[i - 1].HeaderText;
                                worksheet.Cells[startRow, startColumn].Style.Font.Bold = true; // Đặt đậm cho tiêu đề cột
                                worksheet.Cells[startRow, startColumn].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề cột
                                worksheet.Cells[startRow, startColumn].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                                startColumn++;
                            }

                            // Tạo đường viền cho hàng cuối cùng
                            int lastRow = startRow + dataGridView1.Rows.Count;
                            worksheet.Cells[lastRow, 1, lastRow, dataGridView1.Columns.Count].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                            // Xuất dữ liệu từ DataGridView
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                startColumn = 1;
                                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                                {
                                    worksheet.Cells[startRow + 1 + i, startColumn].Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                                    startColumn++;
                                }
                            }

                            // Tạo hàng dọc
                            for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                            {
                                string columnAddress = $"{GetExcelColumnName(i)}{startRow + 1}:{GetExcelColumnName(i)}{lastRow}";
                                worksheet.Cells[columnAddress].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
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
        // Phương thức hỗ trợ để chuyển đổi chỉ số cột thành tên cột Excel
        private string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
