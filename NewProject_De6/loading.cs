using NewProject_De6.FormMain;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewProject_De6
{
    public partial class loading : Form
    {
        private Timer timer;
        public loading()
        {
            InitializeComponent();
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            // Tạo một Timer với thời gian là 5 giây
            timer = new Timer();
            timer.Interval = 5500; // 5 giây
            timer.Tick += Timer_Tick;

            // Bắt đầu đếm thời gian khi Form được hiển thị
            Shown += (sender, e) => timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Dừng Timer
            timer.Stop();

            // Hiển thị Form1
            giaodiendangnhap form1 = new giaodiendangnhap();
            form1.Show();

            this.Hide();
        }

    }
}

