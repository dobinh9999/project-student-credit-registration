using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using NewProject_De6.FormMain;

namespace NewProject_De6
{
    public partial class tracuucongno : Form
    {
        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();
        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            using (SqlConnection connection = new SqlConnection(str))
            {
                try
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();  // Khởi tạo command
                    command.CommandText = "SELECT * FROM tracuucongno";

                    SqlDataAdapter adapter = new SqlDataAdapter();  // Khởi tạo adapter
                    adapter.SelectCommand = command;

                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    dataGridView1.DataSource = table;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
            }
        }
        public tracuucongno()
        {
            InitializeComponent();
        }
        public giaodiendangnhap MainFormReference { get; set; }
        private void tracuucongno_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet1.tracuucongno' table. You can move, or remove it, as needed.
            //this.tracuucongnoTableAdapter.Fill(this.project1DataSet1.tracuucongno);
            connection = new SqlConnection(str); //tạo kết nỗi mới đến chuỗi str
            connection.Open();// mở kết nối khi tải dữ liệu vào form 
            loaddata();// hiện dữ liệu vào form
            connection.Close();//đóng kết nối khi tải dữ liệu xong

            if (MainFormReference != null && MainFormReference.UserRole == "user")
            {
                Console.WriteLine("Disabling buttons for user.");
                // Tắt chức năng thêm và xóa
                button3.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !");
                return;
            }

            tb_macn.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_mahp.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_tenhp.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_stc.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
            tb_sotien.Text = dataGridView1.Rows[i].Cells[4].Value?.ToString();

            dataGridView1.ForeColor = Color.Black;
        }

        private void button2_Click(object sender, EventArgs e)//thêm
        {
            if (string.IsNullOrWhiteSpace(tb_macn.Text) || string.IsNullOrWhiteSpace(tb_mahp.Text) || string.IsNullOrWhiteSpace(tb_tenhp.Text) || string.IsNullOrWhiteSpace(tb_stc.Text) || string.IsNullOrWhiteSpace(tb_sotien.Text))
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
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM tracuucongno WHERE MaCN = @MaCN", connection);
                    checkCommand.Parameters.AddWithValue("@MaCN", tb_macn.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Lỗi : Mã công nợ đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã công nợ khác.", "thông báo !");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO tracuucongno (MaCN, MaHP, TenHP,SoTinChi,Sotien) VALUES ( @MaCN,@MaHP,@TenHP, @SoTinChi,@Sotien)", connection);
                    command.Parameters.AddWithValue("@MaCN", tb_macn.Text);
                    command.Parameters.AddWithValue("@MaHP", tb_mahp.Text);
                    command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
                    command.Parameters.AddWithValue("@SoTinChi", tb_stc.Text);
                    command.Parameters.AddWithValue("@Sotien", tb_sotien.Text);
                    
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

        private void button3_Click(object sender, EventArgs e) // xóa
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (tb_macn.Text != "")
                    {
                        using (SqlConnection connection = new SqlConnection(str))
                        {
                            connection.Open();
                            SqlCommand command = new SqlCommand("DELETE FROM tracuucongno WHERE MaCN = @MaCN", connection);
                            command.Parameters.AddWithValue("@MaCN", tb_macn.Text);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Xóa dữ liệu thành công.", "thông báo !");
                            tb_macn.ResetText();
                            tb_mahp.ResetText();
                            tb_tenhp.ResetText();
                            tb_stc.ResetText();
                            tb_sotien.ResetText();
                        }
                        loaddata(); // Reload data after deleting
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng chọn sinh viên cần xóa.", "thông báo !");
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

        private void button4_Click(object sender, EventArgs e)// sửa
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();
                    // Kiểm tra xem có mã lớp nào khác với mã khoa hiện tại không nếu có là đang sửa mã lớp 
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM tracuucongno WHERE macn = @macn", connection);
                    checkCommand.Parameters.AddWithValue("@macn", tb_macn.Text);

                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE tracuucongno SET mahp = @mahp , TenHP = @TenHP , sotinchi = @sotinchi , sotien = @sotien WHERE macn = @macn", connection);
                        command.Parameters.AddWithValue("@macn", tb_macn.Text);
                        command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                        command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
                        command.Parameters.AddWithValue("@SoTinChi", tb_stc.Text);
                        command.Parameters.AddWithValue("@sotien", tb_sotien.Text);
                        command.ExecuteNonQuery();
                        connection.Close();
                        loaddata();

                        MessageBox.Show("Sửa thông tin thành công.", "thông báo !");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã lớp.", "thông báo !");
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

        private void button6_Click(object sender, EventArgs e) // MOI
        {
            tb_macn.ResetText();
            tb_mahp.ResetText();
            tb_tenhp.ResetText();
            tb_stc.ResetText();
            tb_sotien.ResetText();
            MessageBox.Show("làm mới thành công.", "thông báo !");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            loaddata();
        }

        private void button1_Click(object sender, EventArgs e) // chức năng tìm kiếm 
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

                    command.CommandText = "SELECT * FROM tracuucongno WHERE   macn LIKE @searchText or  mahp LIKE @searchText or TenHP LIKE @searchText or sotinchi  LIKE @searchText or  Sotien LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");
                    
                    
                    SqlDataReader reader = command.ExecuteReader();
                    
                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("macn");
                    dataTable.Columns.Add("mahp");
                    dataTable.Columns.Add("TenHP");
                    dataTable.Columns.Add("sotinchi");
                    dataTable.Columns.Add("Sotien");

                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["macn"] = reader["macn"].ToString();
                        newRow["mahp"] = reader["mahp"].ToString();
                        newRow["TenHP"] = reader["TenHP"].ToString();
                        newRow["sotinchi"] = reader["sotinchi"].ToString();
                        newRow["Sotien"] = reader["Sotien"].ToString();


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

        private void button5_Click(object sender, EventArgs e) // CHUC NANG TIM KIEM
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_DScongno.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_CÔNG NỢ");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Công Nợ";
                            worksheet.Cells["A1:E1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:E2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A3:E3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:E3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:E3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:E1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:E2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:E3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A10:E10"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A10"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:E10"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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

        private void button8_Click(object sender, EventArgs e)
        {
            thanhtoan tt = new thanhtoan();
            tt.ShowDialog();
            
               
            
        }
    }
    
}
