using NewProject_De6.FormMain;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewProject_De6
{
    public partial class HomeForm : Form
    {

        public HomeForm()
        {
            InitializeComponent();
            
            // Gắn sự kiện Load của form
            this.Load += HomeForm_Load;
           
            // Khởi tạo MainFormReference ở đây hoặc tại điểm khởi tạo khác
            MainFormReference = new giaodiendangnhap();

        }
        public giaodiendangnhap MainFormReference { get; set; }

        private void HomeForm_Load(object sender, EventArgs e)
        {

            // Thiết lập kích thước của Form2 thành kích thước tối đa (full screen)
            this.WindowState = FormWindowState.Maximized;


            // Kiểm tra quyền hạn của tài khoản và tắt chức năng 'sinhvien' nếu cần thiết
            if (MainFormReference != null && MainFormReference.UserRole == "user")
            {
               // Tắt chức năng 'sinhvien'
               SinhVien.Enabled = false;
                // Lop.Enabled = false;
                Khoa.Enabled = false;
                Nganh.Enabled = false;
                ThongKe.Enabled = false;
            }
            
        }
        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát không ?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                // Đóng cả ứng dụng
                Application.Exit();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn đăng xuất không ?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Đóng form hiện tại
                this.Close();

                // Mở form đăng nhập
                giaodiendangnhap form1 = new giaodiendangnhap();

                form1.Show();
            }
        }
        private Form currentFormChild; // hiện form con

        private void OpenChildForm(Form childForm)
        {
            if (currentFormChild != null)
            {
                currentFormChild.Close();
            }
            currentFormChild = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            panel_body.Controls.Add(childForm);
            panel_body.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
        }
        
        private void SinhVien_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            TT_SinhVien formTT_SinhVien = new TT_SinhVien();
                            formTT_SinhVien.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formTT_SinhVien);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new TT_SinhVien());
                        }

                        label1.Text = SinhVien.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
        }
        private void Lop_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            TT_lop formTT_Lop = new TT_lop();
                            formTT_Lop.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formTT_Lop);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new TT_lop());
                        }

                        label1.Text = Lop.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
        }



        private void Khoa_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Khoa());
            label1.Text = Khoa.Text;
        }

        private void Nganh_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Nganh());
            label1.Text = Nganh.Text;
        }

        private void PhieuDangKy_Click(object sender, EventArgs e)
        {
            
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            PhieuDK formPhieuDK = new PhieuDK();
                            formPhieuDK.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formPhieuDK);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new PhieuDK());
                        }

                        label1.Text = PhieuDangKy.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
        }

        private void LopHP_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            lopHP formlopHP = new lopHP();
                            formlopHP.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formlopHP);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new lopHP());
                        }

                        label1.Text = LopHP.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
           
        }
        private void TT_MON_DK_Click(object sender, EventArgs e)
        {

            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            TT_mon_hp formTT_mon_hp = new TT_mon_hp();
                            formTT_mon_hp.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formTT_mon_hp);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new Lichhoc());
                        }

                        label1.Text = button1.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }


           
        }
        private void InAn_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            Lichhoc formLichhoc = new Lichhoc();
                            formLichhoc.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formLichhoc);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new Lichhoc());
                        }

                        label1.Text = LichHoc.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
           
        }
        private void ThongKe_Click(object sender, EventArgs e)
        {
            OpenChildForm(new ThongKe());
            label1.Text = ThongKe.Text;
        }
        private void TroGiup_Click(object sender, EventArgs e)
        {
            OpenChildForm(new TroGiup());
            label1.Text = TroGiup.Text;
        }
        private void tracuucongno_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra MainFormReference để đảm bảo nó không phải là null
                if (MainFormReference != null)
                {
                    // Kiểm tra có phải là giaodiendangnhap không
                    if (MainFormReference is giaodiendangnhap)
                    {
                        giaodiendangnhap formReference = (giaodiendangnhap)MainFormReference;

                        // Kiểm tra UserRole
                        if (formReference.UserRole != "admin")
                        {
                            // Nếu đăng nhập bằng user, thực hiện hành động cho user
                            tracuucongno formtracuucongno = new tracuucongno();
                            formtracuucongno.MainFormReference = formReference; // Gán tham chiếu đến form giaodiendangnhap
                            OpenChildForm(formtracuucongno);
                        }
                        if (formReference.UserRole == "admin")
                        {
                            // Nếu đăng nhập bằng admin, thực hiện hành động cho admin
                            OpenChildForm(new tracuucongno());
                        }

                        label1.Text = tracuucongno.Text;
                    }
                    else
                    {
                        MessageBox.Show("MainFormReference không phải là giaodiendangnhap", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("MainFormReference không được khởi tạo", "Lỗi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi");
            }
           
        }
        
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (currentFormChild != null)
            {
                currentFormChild.Close();

            }
            label1.Text = " WINFORM | C# | PHẦN MỀM QUẢN LÝ ĐĂNG KÝ TÍN CHỈ SINH VIÊN ";
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }




        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void qUẢNLÝNGANHToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void fToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked_1(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        
    }
}
