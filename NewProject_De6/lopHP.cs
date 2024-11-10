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
    public partial class lopHP : Form
    {
        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();
        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from qlhp_ttlophp";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }
      
        public lopHP()
        {
            InitializeComponent();
        }

        public giaodiendangnhap MainFormReference { get; set; }
        private void lopHP_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet.qlhp_ttlophp' table. You can move, or remove it, as needed.
            //this.qlhp_ttlophpTableAdapter.Fill(this.project1DataSet.qlhp_ttlophp);
            // TODO: This line of code loads data into the 'project1DataSet.qlhp_ttlophp' table. You can move, or remove it, as needed.
            //this.qlhp_ttlophpTableAdapter.Fill(this.project1DataSet.qlhp_ttlophp);
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
                button6.Enabled = false;
                button5.Enabled = false;
            }
        }
        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !", "thông báo !");
                return;
            }
            tb_mahp.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_tenlophp.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_siso.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();
            tb_tenhp.Text = dataGridView1.Rows[i].Cells[3].Value?.ToString();
        }
        

        private void button1_Click(object sender, EventArgs e)//chức năng thêm 
        {
            if (string.IsNullOrWhiteSpace(tb_tenlophp.Text) || string.IsNullOrWhiteSpace(tb_siso.Text) || string.IsNullOrWhiteSpace(tb_mahp.Text) || string.IsNullOrWhiteSpace(tb_tenhp.Text))
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
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlhp_ttlophp WHERE mahp = @mahp", connection);
                    checkCommand.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Lỗi : Mã học phần đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã học phần khác.", "thông báo !");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO qlhp_ttlophp (mahp, TenLopHP, siso,TenHP) VALUES ( @mahp,@TenLopHP,@siso, @TenHP)", connection);
                    command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    command.Parameters.AddWithValue("@TenLopHP", tb_tenlophp.Text);
                    command.Parameters.AddWithValue("@siso", tb_siso.Text);
                    command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
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

        private void button2_Click(object sender, EventArgs e)// chức năng xóa 
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (tb_mahp.Text != "")
                    {
                        using (SqlConnection connection = new SqlConnection(str))
                        {
                            connection.Open();
                            SqlCommand command = new SqlCommand("DELETE FROM qlhp_ttlophp WHERE mahp = @mahp", connection);
                            command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Xóa dữ liệu thành công.", "thông báo !");
                            tb_mahp.ResetText();
                            tb_tenlophp.ResetText();
                            tb_siso.ResetText();
                            tb_tenhp.ResetText();
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

        private void button3_Click(object sender, EventArgs e)// chức năng sửa 
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();
                    // Kiểm tra xem có mã học phần nào khác với mã học phần hiện tại không nếu có là đang sửa mã học phần 
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlhp_ttlophp WHERE mahp = @mahp", connection);
                    checkCommand.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                    
                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE qlhp_ttlophp SET TenLopHP = @TenLopHP, siso = @siso , TenHP = @TenHP WHERE mahp = @mahp", connection);
                        command.Parameters.AddWithValue("@mahp", tb_mahp.Text);
                        command.Parameters.AddWithValue("@TenLopHP", tb_tenlophp.Text);
                        command.Parameters.AddWithValue("@siso", tb_siso.Text);
                        command.Parameters.AddWithValue("@TenHP", tb_tenhp.Text);
                        command.ExecuteNonQuery();
                        connection.Close();
                        loaddata();

                        MessageBox.Show("Sửa thông tin thành công.", "thông báo !");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã học phần.", "thông báo !");
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

        private void button4_Click(object sender, EventArgs e)// chức năng rest 
        {
            tb_tenlophp.ResetText();
            tb_siso.ResetText();
            tb_mahp.ResetText();
            tb_tenhp.ResetText();
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

                    // Kiểm tra xem searchText có phải là số hay không
                    if (int.TryParse(searchText, out int number))
                    {
                        // Nếu là số, tìm kiếm trong cột siso (sĩ số sinh viên)
                        command.CommandText = "SELECT * FROM qlhp_ttlophp WHERE  siso = @searchText  ";
                        command.Parameters.AddWithValue("@searchText", number);
                    }
                    else
                    {
                        // Nếu không phải là số, tìm kiếm trong cột mahp (mã học phần)
                        command.CommandText = "SELECT * FROM qlhp_ttlophp WHERE   mahp LIKE @searchText or TenLopHP LIKE @searchText or TenHP LIKE @searchText";
                        command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");
                    }
                    
                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("MaHP");
                    dataTable.Columns.Add("TenLopHP");
                    dataTable.Columns.Add("siso");
                    dataTable.Columns.Add("TenHP");

                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["MaHP"] = reader["MaHP"].ToString();
                        newRow["TenLopHP"] = reader["TenLopHP"].ToString();
                        newRow["siso"] = reader["siso"].ToString();
                        newRow["TenHP"] = reader["TenHP"].ToString();

                        // Thêm dòng vào DataTable
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

        private void button7_Click(object sender, EventArgs e)// hiển thị lại dữ liệu 
        {
            loaddata();
        }

        private void button5_Click(object sender, EventArgs e)//chức năng xuất excel
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_LOPHP.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_LỚP HỌC PHẦN");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Lớp Học Phần";
                            worksheet.Cells["A1:D1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:D2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:D3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:D1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:D2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:D3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A22:D22"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A22"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:D22"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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
