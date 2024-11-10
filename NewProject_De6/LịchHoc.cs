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
    public partial class Lichhoc : Form
    {
        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();


        public Lichhoc()
        {
            InitializeComponent();
            
        }

        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from inlich";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }

        public giaodiendangnhap MainFormReference { get; set; }
        private void Lichhoc_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet4.inlich' table. You can move, or remove it, as needed.
            //this.inlichTableAdapter.Fill(this.project1DataSet4.inlich);
           
            // TODO: This line of code loads data into the 'project1DataSet.inlich' table. You can move, or remove it, as needed.
            //this.inlichTableAdapter.Fill(this.project1DataSet.inlich);
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
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }

            tb_malich.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_mahp.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_tenhp.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_tengv.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
            tb_tg.Text = dataGridView1.Rows[i].Cells[4].Value?.ToString();
            tb_diadiem.Text = dataGridView1.Rows[i].Cells[5].Value?.ToString();
        }

        private void button2_Click(object sender, EventArgs e)// chức năng thêm
        {
            if (string.IsNullOrWhiteSpace(tb_malich.Text) || string.IsNullOrWhiteSpace(tb_tenhp.Text) || string.IsNullOrWhiteSpace(tb_tengv.Text) || string.IsNullOrWhiteSpace(tb_tg.Text) || string.IsNullOrWhiteSpace(tb_diadiem.Text) || string.IsNullOrWhiteSpace(tb_mahp.Text))
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
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM inlich WHERE malich = @malich", connection);
                    checkCommand.Parameters.AddWithValue("@malich", tb_malich.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Lỗi : Mã lịch học đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã lịch khác.", "thông báo !");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO inlich (malich,mahp, TeHP, TenGV,time,diadiem) VALUES ( @malich,@mahp,@TeHP,@TenGV, @time,@diadiem)", connection);
                    command.Parameters.AddWithValue("@malich", tb_malich.Text);
                    command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    command.Parameters.AddWithValue("@TeHP", tb_tenhp.Text);
                    command.Parameters.AddWithValue("@TenGV", tb_tengv.Text);
                    command.Parameters.AddWithValue("@time", tb_tg.Text);
                    command.Parameters.AddWithValue("@diadiem", tb_diadiem.Text);
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

        private void button3_Click(object sender, EventArgs e)// xoa
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (tb_malich.Text != "")
                    {
                        using (SqlConnection connection = new SqlConnection(str))
                        {
                            connection.Open();
                            SqlCommand command = new SqlCommand("DELETE FROM inlich WHERE malich = @malich", connection);
                            command.Parameters.AddWithValue("@malich", tb_malich.Text);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Xóa dữ liệu thành công.", "thông báo !");
                            tb_malich.ResetText();
                            tb_mahp.ResetText();
                            tb_tenhp.ResetText();
                            tb_tengv.ResetText();
                            tb_tg.ResetText();
                            tb_diadiem.ResetText();
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

        private void button4_Click(object sender, EventArgs e)// sua
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();
                    // Kiểm tra xem có mã lớp nào khác với mã khoa hiện tại không nếu có là đang sửa mã lớp 
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM inlich WHERE malich = @malich", connection);
                    checkCommand.Parameters.AddWithValue("@malich", tb_malich.Text);

                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE inlich SET mahp = @mahp,TeHP = @TeHP, TenGV = @TenGV , time = @time, diadiem = @diadiem WHERE malich = @malich", connection);
                        command.Parameters.AddWithValue("@malich", tb_malich.Text);
                        command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                        command.Parameters.AddWithValue("@TeHP", tb_tenhp.Text);
                        command.Parameters.AddWithValue("@TenGV", tb_tengv.Text);
                        command.Parameters.AddWithValue("@time", tb_tg.Text);
                        command.Parameters.AddWithValue("@diadiem", tb_diadiem.Text);
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

        private void button6_Click(object sender, EventArgs e)
        {
            tb_malich.ResetText();
            tb_tenhp.ResetText();
            tb_tengv.ResetText();
            tb_tg.ResetText();
            tb_diadiem.ResetText();
            MessageBox.Show("làm mới thành công.", "thông báo !");
        }

        private void button5_Click(object sender, EventArgs e)// ecxel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_lichhoc.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_Lịch Học");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Lịch Học";
                            worksheet.Cells["A2"].Value = "Thông tin chi tiết";
                            worksheet.Cells["A1:F1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:F2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 15; // Đặt kích thước font cho tiêu đề
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
                            worksheet.Cells["A13:F13"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A13"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:F13"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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

        private void button1_Click(object sender, EventArgs e)// tìm kiếm   
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
                    command.CommandText = "SELECT * FROM inlich WHERE malich LIKE @searchText OR tehp LIKE @searchText OR TenGV LIKE @searchText OR time LIKE @searchText OR diadiem LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("malich");
                    dataTable.Columns.Add("TeHP");
                    dataTable.Columns.Add("TenGV");
                    dataTable.Columns.Add("time");
                    dataTable.Columns.Add("diadiem");

                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["malich"] = reader["malich"].ToString();
                        newRow["TeHP"] = reader["TeHP"].ToString();
                        newRow["TenGV"] = reader["TenGV"].ToString();
                        newRow["time"] = reader["time"].ToString();
                        newRow["diadiem"] = reader["diadiem"].ToString();


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

        private void button7_Click(object sender, EventArgs e)
        {
            loaddata();
        }
    }
}
