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
    public partial class DangKy : Form
    {
        Module.DataAccess dtBase = new Module.DataAccess();
        public DangKy()
        {
            InitializeComponent();
        }
        private void txtSDT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if((Convert.ToInt16(e.KeyChar) < Convert.ToInt16('0') || Convert.ToInt16(e.KeyChar) > Convert.ToInt16('9')) && 
                Convert.ToInt16(e.KeyChar) != 8) {
                MessageBox.Show("Bạn chỉ được nhập số nguyên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Handled = true;
            }
        }
        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (txtMaSV.Text.Trim() == "" || txtHoTen.Text.Trim() == "" || txtPass.Text.Trim() == "" || txtSDT.Text.Trim() == ""
                || txtEmail.Text.Trim() == "" || cmbGioiTinh.Text.Trim() == "" || txtDiaChi.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin để tạo tài khoản!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaSV.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from SINHVIEN where MASV='" + txtMaSV.Text + "'");
                if (tbDanhSach.Rows.Count > 0)
                {
                    MessageBox.Show("Danh sách đã có mã " + txtMaSV.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSV.Focus();
                    return;
                }
                if(txtMaSV.Text.StartsWith("18120") == false && txtMaSV.Text.StartsWith("18121") == false && txtMaSV.Text.StartsWith("19120") == false && txtMaSV.Text.StartsWith("19121") == false 
                    && txtMaSV.Text.StartsWith("20120") == false && txtMaSV.Text.StartsWith("20121") == false && txtMaSV.Text.StartsWith("21120") == false && txtMaSV.Text.StartsWith("21121") == false)
                {
                    MessageBox.Show("Mã sinh viên không đúng định dạng!\nVui lòng nhập lại mã sinh viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSV.Focus();
                    return;
                }
                if (txtMaSV.Text.Length != 9)
                {
                    MessageBox.Show("Mã sinh viên cần có đủ 9 ký tự!Vui lòng nhập đúng định dạng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSV.Focus();
                    return;
                }
                if (int.Parse(DateTime.Now.ToString().Substring(6, 5)) - int.Parse(dtpNgaySinh.Value.ToString().Substring(6, 5)) < 18)
                {
                    MessageBox.Show("Ngày sinh không đúng!\nVui lòng nhập lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dtpNgaySinh.Focus();
                    return;
                }
                if (txtPass.Text.Length < 8)
                {
                    MessageBox.Show("Mật khẩu cần có từ 8 ký tự trở lên '"+ dtpNgaySinh.Value + "' '"+DateTime.Now+"'!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtPass.Focus();
                    return;
                }
                if (txtSDT.Text.Length != 10 && txtSDT.Text.Length != 11)
                {
                    MessageBox.Show("SĐT không đúng định dạng!\nVui lòng nhập lại!\n(SĐT phải có đủ 10 hoặc 11 chữ số)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtSDT.Focus();
                    return;
                }
                if (txtSDT.Text.StartsWith("84") == false && txtSDT.Text.StartsWith("0") == false)
                {
                    MessageBox.Show("SĐT không đúng định dạng!\nVui lòng nhập lại!\n(SĐT phải đầu số là (84) hoặc (0)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtSDT.Focus();
                    return;
                }
                if (txtSDT.Text.StartsWith("0") && txtSDT.Text.Length != 10)
                {
                    MessageBox.Show("SDT không đúng định dạng. SĐT có đầu số là 0 thì cần có đủ 10 chữ số", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtSDT.Focus();
                    return;
                }
                if (txtSDT.Text.StartsWith("84") && txtSDT.Text.Length != 11)
                {
                    MessageBox.Show("SDT không đúng định dạng. SĐT có đầu số là (84) thì cần có đủ 11 chữ số", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtSDT.Focus();
                    return;
                }

                if (txtEmail.ToString().Contains("@gmail.com") != true && txtEmail.ToString().Contains("@st.utc.edu.vn") != true)
                {
                    MessageBox.Show("Email không đúng định dạng!\nVui lòng nhập đúng email!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtEmail.Focus();
                    return;
                }
                string sqlInsert = "insert into SINHVIEN values(N'" + txtMaSV.Text + "', N'" + txtPass.Text + "', N'" + txtHoTen.Text + "', N'" + dtpNgaySinh.Value.ToString("yyyy-MM-dd") + "', N'" + cmbGioiTinh.Text + "', N'" + txtDiaChi.Text + "', N'" + txtSDT.Text + "', N'" + txtEmail.Text + "')";
                dtBase.Updatedate(sqlInsert);

                if(MessageBox.Show("Bạn đã đăng ký tài khoản thành công!\nĐăng nhập để sử dụng ứng dụng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    this.Close();
;               }
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
