using NewProject_De6.FormMain;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace NewProject_De6
{
    public partial class TT_lop : Form
    {

        SqlConnection connection;
        SqlCommand command;
        string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();
        void loaddata()//thực hiện truyền tải dữ liệu vào form
        {
            command = connection.CreateCommand();
            command.CommandText = "select * from qlsv_lopcuasv";
            // command.ExecuteNonQuery(); // thực thi câu truy vấn
            adapter.SelectCommand = command;
            table.Clear();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }
        public giaodiendangnhap MainFormReference { get; set; }
        private void TT_lop_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'project1DataSet.qlsv_lopcuasv' table. You can move, or remove it, as needed.
            //this.qlsv_lopcuasvTableAdapter.Fill(this.project1DataSet.qlsv_lopcuasv);
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
                button6.Enabled = false;
            }
        }
       
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)// hiện thông tin trên textbox
        {
            int i = e.RowIndex;
            if (i < 0)
            {
                MessageBox.Show("Vui lòng chọn lại !");
                return;
            }

            tb_malop.Text = dataGridView1.Rows[i].Cells[0].Value?.ToString();
            tb_tenlop.Text = dataGridView1.Rows[i].Cells[1].Value?.ToString();
            tb_tencv.Text = dataGridView1.Rows[i].Cells[2].Value?.ToString();

            dataGridView1.ForeColor = Color.Black;
        }

        private void button5_Click(object sender, EventArgs e)// chức năng làm mới 
        {
            tb_malop.ResetText();
            tb_tenlop.ResetText();
            tb_tencv.ResetText();

            MessageBox.Show("làm mới thành công.");
        }

        private void button1_Click(object sender, EventArgs e)// chức năng thêm 
        {
            if (string.IsNullOrWhiteSpace(tb_malop.Text) || string.IsNullOrWhiteSpace(tb_tenlop.Text) || string.IsNullOrWhiteSpace(tb_tencv.Text))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin.");
                return;
            }

            using (SqlConnection connection = new SqlConnection(str))
            {
                try
                {

                    connection.Open();

                    // Check malop
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_lopcuasv WHERE malop = @malop", connection);
                    checkCommand.Parameters.AddWithValue("@malop", tb_malop.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBox.Show("Mã lớp đã tồn tại trong cơ sở dữ liệu. Vui lòng chọn mã lớp khác.");
                        return;
                    }

                    SqlCommand command = new SqlCommand("INSERT INTO qlsv_lopcuasv (malop, tenlop, tencv) VALUES (@malop, @tenlop, @tencv)", connection);
                    command.Parameters.AddWithValue("@malop", tb_malop.Text);
                    command.Parameters.AddWithValue("@tenlop", tb_tenlop.Text);
                    command.Parameters.AddWithValue("@tencv", tb_tencv.Text);
                    int rowsAffected = command.ExecuteNonQuery();
                    
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Thêm thành công.");
                        loaddata(); // Gọi hàm loaddata() để cập nhật DataGridView
                    }
                    else
                    {
                        MessageBox.Show("Thêm không thành công. Vui lòng kiểm tra lại dữ liệu.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                finally
                {
                    connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
                }
               
            }
        }
        private void button2_Click(object sender, EventArgs e)// chức năng sửa 
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(str))
                {
                    connection.Open();

                    // Kiểm tra xem có mã lớp nào khác với mã lớp hiện tại không nếu có là đang sửa mã lớp 
                    SqlCommand checkCommand = new SqlCommand("SELECT COUNT(*) FROM qlsv_lopcuasv WHERE malop = @malop", connection);
                    checkCommand.Parameters.AddWithValue("@malop", tb_malop.Text);
                    int existingRecords = (int)checkCommand.ExecuteScalar();

                    if (existingRecords > 0)
                    {
                        SqlCommand command = new SqlCommand("UPDATE qlsv_lopcuasv SET tenlop = @tenlop, tencv = @tencv WHERE malop = @malop", connection);
                        command.Parameters.AddWithValue("@malop", tb_malop.Text);
                        command.Parameters.AddWithValue("@tenlop", tb_tenlop.Text);
                        command.Parameters.AddWithValue("@tencv", tb_tencv.Text);
                        command.ExecuteNonQuery();
                        connection.Close();
                        loaddata();

                        MessageBox.Show("Sửa thông tin thành công.");
                    }
                    else
                    {
                        // hiển thị thông báo lỗi
                        MessageBox.Show("Không được sửa mã lớp.");
                    }
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
            }
        }
        private void button3_Click(object sender, EventArgs e)// chức năng xóa
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa dữ liệu này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(str))
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand("DELETE FROM qlsv_lopcuasv WHERE malop = @malop", connection);
                        command.Parameters.AddWithValue("@malop", tb_malop.Text);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Xóa dữ liệu thành công.");
                        tb_malop.ResetText();
                        tb_tenlop.ResetText();
                        tb_tencv.ResetText();
                    }
                    loaddata(); // Reload data after deleting
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
                finally
                {
                    connection.Close(); // Đảm bảo kết nối được đóng ngay cả khi xảy ra ngoại lệ
                }
            }
        }
        private void button6_Click(object sender, EventArgs e)// chức năng tìm kiếm 
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
                    command.CommandText = "SELECT * FROM qlsv_lopcuasv WHERE malop LIKE @searchText OR tenlop LIKE @searchText OR tencv LIKE @searchText";
                    command.Parameters.AddWithValue("@searchText", "%" + searchText + "%");

                    SqlDataReader reader = command.ExecuteReader();

                    // Tạo cấu trúc cho DataTable (tên cột phải trùng với tên trường trong cơ sở dữ liệu)
                    dataTable.Columns.Add("malop");
                    dataTable.Columns.Add("tenlop");
                    dataTable.Columns.Add("tencv");
                    
                    // Thêm các cột khác tương tự


                    while (reader.Read())
                    {
                        // Thêm dòng vào DataTable
                        DataRow newRow = dataTable.NewRow();
                        newRow["malop"] = reader["malop"].ToString();
                        newRow["tenlop"] = reader["tenlop"].ToString();
                        newRow["tencv"] = reader["tencv"].ToString();
                        

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
        private void button7_Click(object sender, EventArgs e)// chức năng hiển thị lại dữ liệu 
        {
            loaddata();
        }
        private void button4_Click(object sender, EventArgs e)// CHỨC NĂNG XUẤT EXCEL
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_lOP.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("DS_LỚP");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "Danh Sách Lớp";
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
                            worksheet.Cells["A11:C11"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang

                            // tạo hàng dọc
                            worksheet.Cells["A1:A11"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:C11"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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
        public TT_lop()
        {
            InitializeComponent();
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        
    }





}
