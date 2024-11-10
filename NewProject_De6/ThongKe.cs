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

namespace NewProject_De6
{
    public partial class ThongKe : Form
    {
        SqlConnection connection;
        SqlCommand command;
        SqlDataAdapter adapter = new SqlDataAdapter();
        DataTable table = new DataTable();
        
        
        public ThongKe()
        {
            InitializeComponent();
            string str = @"Data Source=LAPTOP-D90-DOXU\SQLEXPRESS;Initial Catalog=project1;Integrated Security=True";
            connection = new SqlConnection(str);
            command = new SqlCommand();
            adapter = new SqlDataAdapter();
            table = new DataTable();
            // Ẩn CheckBox khi form được tạo
            checkBox1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
        }

        private void ThongKe_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("Sinh viên đăng ký thành công");
            comboBox1.Items.Add("Sinh viên phải đăng ký lại");
            comboBox1.Items.Add("số lượng ngành trong các khoa");
            comboBox1.Items.Add("số lượng sinh viên theo giới tính");           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedItem.ToString() == "Sinh viên đăng ký thành công" )
            {
                checkBox1.Visible = true;
                button2.Visible = true;
                button3.Visible = false;
                button4.Visible = false;
            }
            else if(comboBox1.SelectedItem.ToString() == "Sinh viên phải đăng ký lại")
            {
                checkBox1.Visible = true;
                button3.Visible = true;
                button2.Visible = false;
                button4.Visible = false;
            }
            else if (comboBox1.SelectedItem.ToString() == "số lượng ngành trong các khoa")
            {
                checkBox1.Visible = true;
                button3.Visible = false;
                button2.Visible = false;
                button4.Visible = true;
            }
            else if (comboBox1.SelectedItem.ToString() == "số lượng sinh viên theo giới tính")
            {
                checkBox1.Visible = true;
                button3.Visible = false;
                button2.Visible = false;
                button4.Visible = false;
            }
            
            else
            {
                checkBox1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
            }
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(comboBox1.Text))
            {
                MessageBox.Show("Vui lòng nhập thông tin.", "thông báo !");
                return;
            }

            string selectedOption = comboBox1.SelectedItem.ToString();
            DataTable dataTable = new DataTable();
            
            try
            {
                connection.Open();
                command = connection.CreateCommand();

                /*if (selectedOption == "Sinh viên đăng ký thành công")
                {
                    checkBox1.Visible = true;
                    command.CommandText = "SELECT COUNT(s.MSV) AS 'Số Lượng SV Phải đăng ký lại'FROM qlsv_sinhvien s LEFT JOIN qlhp_phieudkHP p ON s.MSV = p.MSV WHERE p.MSV IS NULL";
                }
                else
                {
                    // Ẩn CheckBox nếu không phải là mục cần hiển thị CheckBox
                    checkBox1.Visible = false;
                }*/

                if (selectedOption == "Sinh viên đăng ký thành công")
                {
                    if(checkBox1.Checked)
                    {
                        command.CommandText = "select * from qlhp_phieudkHP";
                    }
                    else
                    {
                        command.CommandText = "SELECT COUNT(*) AS 'Số SV đăng ký thành công'  FROM [dbo].[qlhp_phieudkHP]";
                    }
                }
                 if (selectedOption == "Sinh viên phải đăng ký lại")
                {
                    if (checkBox1.Checked)
                    {
                        command.CommandText = "SELECT s.* FROM qlsv_sinhvien s LEFT JOIN qlhp_phieudkHP p ON s.MSV = p.MSV WHERE p.MSV IS NULL;";
                    }
                    else
                    {
                        command.CommandText = "SELECT COUNT(s.MSV) AS 'Số Lượng SV Phải đăng ký lại'FROM qlsv_sinhvien s LEFT JOIN qlhp_phieudkHP p ON s.MSV = p.MSV WHERE p.MSV IS NULL";
                    }
                    
                }
                if (selectedOption == "số lượng ngành trong các khoa")
                {
                    if (checkBox1.Checked)
                    {
                        command.CommandText = "SELECT k.makhoa AS'Mã Khoa', k.tenkhoa AS 'Tên khoa', n.manganh AS'Mã Ngành',n.tennganh AS'Tên Ngành'  \r\nFROM qlsv_nganh  n join qlsv_khoa k on n.makhoa = k.makhoa\r\nGROUP BY k.makhoa,k.tenkhoa,n.manganh,n.tennganh ";
                        
                    }
                    else
                    {
                        command.CommandText = "SELECT \r\n\t   k.makhoa AS 'Mã khoa',\r\n\t   k.tenkhoa AS 'Tên Khoa',\r\n\t   count(n.manganh)  AS 'số lượng ngành trong Một khoa'\r\nFROM \r\n\tqlsv_khoa k\r\n\tLEFT JOIN \r\n\tqlsv_nganh n on k.makhoa = n.makhoa\r\nGROUP BY \r\n\tk.makhoa,k.tenkhoa";
                    }

                }
                if (selectedOption == "số lượng sinh viên theo giới tính")
                {
                    if (checkBox1.Checked)
                    {
                        command.CommandText = "SELECT *   FROM qlsv_sinhvien where GioiTinh = 'Nam'";

                    }
                    else
                    {
                        command.CommandText = "SELECT GioiTinh AS 'GIỚI TÍNH', COUNT(*) AS 'SỐ LƯỢNG' FROM qlsv_sinhvien GROUP BY GioiTinh";
                    }

                }
               


                adapter.SelectCommand = command;
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo lỗi");
            }
            finally
            {
                connection.Close();
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_BaoCaoThongKe.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("BÁO CÁO THỐNG KÊ");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "BÁO CÁO DANH SÁCH ĐÃ THỐNG KÊ";
                            worksheet.Cells["A2"].Value = "Sinh Viên Đăng Ký Thành Công";
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_BaoCaoThongKe.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("BÁO CÁO THỐNG KÊ");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "BÁO CÁO DANH SÁCH ĐÃ THỐNG KÊ";
                            worksheet.Cells["A2"].Value = "Sinh Viên Đăng Ký Lại ";
                            worksheet.Cells["A1"].Style.Font.Size = 25; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 20; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1:I1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:I2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A3:I3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:I3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:I3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề         
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề
                            worksheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:I1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:I2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:I3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang                          

                            // tạo hàng dọc
                            worksheet.Cells["A1:A3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:I3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    saveFileDialog.FileName = "output_BaoCaoThongKe.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("BÁO CÁO THỐNG KÊ");

                            // Thêm tiêu đề "Danh sách sinh viên"
                            worksheet.Cells["A1"].Value = "BÁO CÁO DANH SÁCH ĐÃ THỐNG KÊ";
                            worksheet.Cells["A2"].Value = "Thông tin chi tiết";
                            worksheet.Cells["A1"].Style.Font.Size = 15; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A2"].Style.Font.Size = 10; // Đặt kích thước font cho tiêu đề
                            worksheet.Cells["A1:C1"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A2:C2"].Merge = true; // Merge các ô từ A1 đến D1
                            worksheet.Cells["A3:C3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A3:C3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["A3:C3"].Style.Font.Size = 12; // Đặt kích thước font cho tiêu đề         
                            worksheet.Cells["A1"].Style.Font.Bold = true; // Đặt đậm cho tiêu đề
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề
                            worksheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Căn giữa tiêu đề

                            // Tạo hàng ngang
                            worksheet.Cells["A1:C1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A2:C2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang
                            worksheet.Cells["A3:C3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng ngang                          

                            // tạo hàng dọc
                            worksheet.Cells["A1:A3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc
                            worksheet.Cells["A1:C3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; // Đặt kiểu đường viền cho hàng dọc

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
    }
    
}
