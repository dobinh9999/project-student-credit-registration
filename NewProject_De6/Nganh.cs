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
    public partial class Nganh : Form
    {
        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();
        public Nganh()
        {
            InitializeComponent();
        }
        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from qlsv_nganh";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }
        private void Nganh_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet3.qlsv_nganh' table. You can move, or remove it, as needed.
            //this.qlsv_nganhTableAdapter.Fill(this.project1DataSet3.qlsv_nganh);
            // TODO: This line of code loads data into the 'project1DataSet.qlsv_nganh' table. You can move, or remove it, as needed.
            //this.qlsv_nganhTableAdapter.Fill(this.project1DataSet.qlsv_nganh);
            connection = new SqlConnection(str); //tạo kết nỗi mới đến chuỗi str
            connection.Open();// mở kết nối khi tải dữ liệu vào form 
            loaddata();// hiện dữ liệu vào form
            connection.Close();//đóng kết nối khi tải dữ liệu xong
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }

            tb_manganh.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_tennganh.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_makhoa.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_mahp.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
        }

        private void button1_Click_1(object sender, EventArgs e)//chức năng thêm
        {
            if (string.IsNullOrWhiteSpace(tb_manganh.Text) || string.IsNullOrWhiteSpace(tb_tennganh.Text) || string.IsNullOrWhiteSpace(tb_makhoa.Text) || string.IsNullOrWhiteSpace(tb_mahp.Text))
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
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_nganh WHERE manganh = @manganh", connection);
                    checkCommand.Parameters.AddWithValue("@manganh", tb_manganh.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Mã ngành đã tồn tại trong cơ sở dữ liệu.\n Vui lòng chọn mã ngành khác.", "thông báo !");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO qlsv_nganh (manganh, tennganh, makhoa,mahp) VALUES (@manganh, @tennganh, @makhoa,@mahp)", connection);
                    command.Parameters.AddWithValue("@manganh", tb_manganh.Text);
                    command.Parameters.AddWithValue("@tennganh", tb_tennganh.Text);
                    command.Parameters.AddWithValue("@makhoa", tb_makhoa.Text);
                    command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    int rowsAffected = command.ExecuteNonQuery();

                    
                       
                        loaddata(); // Gọi hàm loaddata() để cập nhật DataGridView

                    MessageBox.Show("Thêm thành công.", "thông báo !");
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

        private void button2_Click(object sender, EventArgs e)// chức năng xóa
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(str))
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand("DELETE FROM qlsv_nganh WHERE manganh = @manganh", connection);
                        command.Parameters.AddWithValue("@manganh", tb_manganh.Text);
                        command.ExecuteNonQuery();
                        
                        tb_manganh.ResetText();
                        tb_tennganh.ResetText();
                        tb_makhoa.ResetText();
                        tb_mahp.ResetText();

                        MessageBox.Show("Xóa dữ liệu thành công.", "thông báo !");
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
        private void button3_Click(object sender, EventArgs e)// chức năng sửa
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();
                    // Kiểm tra xem có mã nganh nào khác với mã khoa hiện tại không nếu có là đang sửa mã nganh 
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_nganh WHERE manganh = @manganh", connection);
                    checkCommand.Parameters.AddWithValue("@manganh", tb_manganh.Text);
                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE qlsv_nganh SET tennganh = @tennganh, makhoa = @makhoa,mahp = @mahp WHERE manganh = @manganh", connection);
                        command.Parameters.AddWithValue("@manganh", tb_manganh.Text);
                        command.Parameters.AddWithValue("@tennganh", tb_tennganh.Text);
                        command.Parameters.AddWithValue("@makhoa", tb_makhoa.Text);
                        command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                        command.ExecuteNonQuery();
                        connection.Close();
                        loaddata();

                        MessageBox.Show("Sửa thông tin thành công.", "thông báo !");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã ngành.", "thông báo !");
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

        private void button4_Click(object sender, EventArgs e)// chức năng reset
        {
            tb_manganh.ResetText();
            tb_tennganh.ResetText();
            tb_makhoa.ResetText();
            tb_mahp.ResetText();
            MessageBox.Show("làm mới thành công.", "thông báo !");
        }

        private void button6_Click(object sender, EventArgs e)//chức năng tìm kiếm 
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
                    command.CommandText = "SELECT * FROM qlsv_nganh WHERE manganh LIKE @searchText OR tennganh LIKE @searchText OR makhoa LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("MaNganh");
                    dataTable.Columns.Add("TenNganh");
                    dataTable.Columns.Add("Makhoa");

                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["MaNganh"] = reader["MaNganh"].ToString();
                        newRow["TenNganh"] = reader["TenNganh"].ToString();
                        newRow["Makhoa"] = reader["Makhoa"].ToString();


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
                MessageBox.Show("Vui lòng nhập mã ngành để tìm kiếm.", "Thông báo!");
            }
        }   

        private void button7_Click(object sender, EventArgs e)// hiển lại dữ liệu
        {
            loaddata();
        }

        private void button5_Click(object sender, EventArgs e)// chức năng in xuất excel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_NGANH.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_NGÀNH");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Ngành";
                            worksheet.Cells["A1:C1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:C2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A3:C3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:C3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:C3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:C1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:C2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:C3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A23:C23"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A23"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:C23"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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
    }
}
