using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BTL_ThuVien
{
    public partial class Form1 : Form
    {
        Module.DataAccess dtBase = new Module.DataAccess();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            if(txtUserName.Text == "" || txtPass.Text == "")
            {
                MessageBox.Show("Bạn cần nhập đầy đủ Username và PassWord", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if(rdbNhanVien.Checked == false && rdbSV.Checked == false)
            {
                MessageBox.Show("Bạn cần chọn chức vụ (Nhân viên hoặc Sinh Viên)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if((txtUserName.Text.ToUpper() == "ADMIN") && txtPass.Text == "12345" && rdbNhanVien.Checked == true)
            {
                MessageBox.Show("Bạn đã đăng nhập thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                frmNVQuanLy NVQuanLy = new frmNVQuanLy();
                NVQuanLy.Show();
            }
            else if(rdbSV.Checked == true)
            {
                DataTable tbTaiKhoan = dtBase.DataSelect("select * from SINHVIEN where MASV='" + txtUserName.Text + "' and PASS = '"+ txtPass.Text +"'");
                if (tbTaiKhoan.Rows.Count == 1)
                {
                    MessageBox.Show("Bạn đã đăng nhập thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    frmSinhVien SV = new frmSinhVien(txtUserName.Text, txtPass.Text);
                    SV.Show();
                }
                else
                {
                    MessageBox.Show("UserName hoặc Password sai! Hoặc tài khoản không tồn tại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }               
            }
            else
            {
                MessageBox.Show("UserName hoặc Password sai!\nVui lòng nhập đúng thông tin tài khoản của bạn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                this.Close();
            } 
        }

        private void lblDKy_Click(object sender, EventArgs e)
        {
            DangKy dk = new DangKy();
            dk.Show();
        }
    }
}
