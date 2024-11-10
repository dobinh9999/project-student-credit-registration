using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewProject_De6.FormMain
{
    public partial class giaodiendangnhap : Form
    {

        private ErrorProvider errorProvider1;

        
        public giaodiendangnhap()
        {
            InitializeComponent();
            matkhau.UseSystemPasswordChar = true;
            // Trong nơi tạo đối tượng của HomeForm

            errorProvider1 = new ErrorProvider();

            
        }
        private void CenterFormOnScreen()// tính toán tìm vị trí để đặt form khi khởi động
        {
            // Tính toán vị trí mới cho form
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            int formWidth = this.Width;
            int formHeight = this.Height;

            int newLeft = (screenWidth - formWidth) / 2;
            int newTop = (screenHeight - formHeight) / 2;

            // Đặt vị trí mới cho form
            this.Location = new System.Drawing.Point(newLeft, newTop);
        }
        public class Account
        {
           
            public string Username { get; set; }
            public string Password { get; set; }
            public string Role { get; set; } // Ví dụ: 'admin' hoặc 'user'
            
        }
        public string UserRole { get; private set; }
        List<Account> accounts = new List<Account>
        {
            new Account { Username = "quanly", Password = "123", Role = "admin" },
            new Account { Username = "sinhvien", Password = "123", Role = "user" },
            // Thêm các tài khoản khác nếu cần thiết
        };
        private void button1_Click(object sender, EventArgs e)
        {
            string username = ten.Text;
            string password = matkhau.Text;
            errorProvider1.Clear();
            if (ten.Text == "")
            {
                errorProvider1.SetError(ten, "Vui lòng nhập tên đăng nhập!");
            }
            else if (matkhau.Text == "")
            {
                errorProvider1.SetError(matkhau, "Vui lòng nhập mật khẩu!");
            }

            
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin", "Thông báo!");
            }
            else
            {
                // tìm kiếm tài khoản từ danh sách 
                Account matchedAccount = accounts.FirstOrDefault(acc => acc.Username == username && acc.Password == password);

                if (matchedAccount != null)
                {
                    

                    // Lấy quyền hạn của tài khoản
                    string role = matchedAccount.Role;

                    // Kiểm tra quyền hạn và thực hiện các hành động tương ứng
                    if (role == "admin")
                    {

                        MessageBox.Show("Đăng nhập thành công", "ADMIN");
                       
                        // Thực hiện hành động cho quyền admin
                        HomeForm form2 = new HomeForm();
                        form2.Show();
                    }

                    else if (role == "user")
                    {
                        MessageBox.Show("Đăng nhập thành công", "USER");

                        UserRole = matchedAccount.Role;               
                        // Mở HomeForm
                        HomeForm form2 = new HomeForm();
                        form2.MainFormReference = this; // Truyền tham chiếu đến MainForm
                        form2.Show();
                    }
                    this.Hide();
                    // Làm sạch sau khi đăng nhập
                    ten.Text = "";
                    matkhau.Text = "";               
                }
                else
                {
                    MessageBox.Show("Tài khoản hoặc mật khẩu không chính xác , vui lòng nhập lại!", "Thông báo!");
                }
            }
        }
        
        private void checkpass_CheckedChanged(object sender, EventArgs e)
        {
            if (checkpass.Checked)
            {
                // Hiển thị mật khẩu
                matkhau.UseSystemPasswordChar = false;
            }
            else
            {
                // Ẩn mật khẩu
                matkhau.UseSystemPasswordChar = true ; 
            }
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show(" Liên hệ admin bình bờ rồ bằng thông tin liên hệ dưới này nhé! \n Phone : 0123456789. \n Gmail : binhking@gmail.com. \n 'cảnh báo' => admin là thánh bịp  =)) ","Thông báo!");
        }
        private void matkhau_TextChanged(object sender, EventArgs e)
        {

        }

        private void giaodiendangnhap_Load(object sender, EventArgs e)
        {
            // Đặt vị trí khởi đầu của form ở giữa màn hình desktop
            CenterFormOnScreen();
        }
        

        private void taikhoan_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
