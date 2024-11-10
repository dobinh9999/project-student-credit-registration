using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewProject_De6
{
    internal class ds_taikhoan
    {
        public class Account
        {
            public string Username { get; set; }
            public string Password { get; set; }
            public string Role { get; set; } // Ví dụ: 'admin' hoặc 'user'
        }
        List<Account> accounts = new List<Account>
        {
            new Account { Username = "binh", Password = "123", Role = "admin" },
            new Account { Username = "user1", Password = "password1", Role = "user" },
            // Thêm các tài khoản khác nếu cần thiết
        };
    }
}
