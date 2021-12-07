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
    public partial class frmSinhVien : Form
    {
        Module.DataAccess dtBase = new Module.DataAccess();

        string username, passuser;
        public frmSinhVien()
        {
            InitializeComponent();
        }
        public frmSinhVien(string name, string pass)
        {
            InitializeComponent();
            username = name;
            passuser = pass;
        }
        private void btnThoatTT_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
            {
                this.Close();
            }
        }

        void LoadDataSach()
        {
            dgvSach.DataSource = dtBase.DataSelect("Select SACH.MASACH, SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, SACH.GIABIA " +
            "from ((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB");
        }

        void LoadDataPhieuMuon()
        {
            dgvMuon.DataSource = dtBase.DataSelect("select PHIEUMUON.MAPHIEUMUON, SACH.TENSACH, PHIEUMUON.SOLUONGMUON, PHIEUMUON.NGAYMUON, PHIEUMUON.NGAYHENTRA, PHIEUMUON.TRANGTHAI " +
                "from PHIEUMUON join SACH on PHIEUMUON.MASACH = SACH.MASACH where MASV='" + username.ToString() + "'");
        }

        void LoadDataPhieuTra()
        {
            dgvPhieuTra.DataSource = dtBase.DataSelect("Select PHIEUTRA.MAPHIEUTRA, PHIEUTRA.MAPHIEUMUON, SACH.TENSACH, PHIEUTRA.NGAYTRA, VIPHAM.ND_VIPHAM, VIPHAM.ND_PHAT, PHIEUTRA.PHINOPPHAT, GHICHU" +
                " from ((PHIEUTRA join PHIEUMUON ON PHIEUTRA.MAPHIEUMUON = PHIEUMUON.MAPHIEUMUON) join SACH ON PHIEUMUON.MASACH = SACH.MASACH ) " +
                "FULL JOIN VIPHAM ON VIPHAM.MAVIPHAM = PHIEUTRA.MAVIPHAM where MASV='" + username.ToString() + "'");
        }

        void loadDataThongBao()
        {
            ///
            string thongbao = "";
            DataTable tbThongBao = dtBase.DataSelect("select * from PHIEUMUON join SACH on PHIEUMUON.MASACH = SACH.MASACH where PHIEUMUON.TRANGTHAI = N'Quá hạn' and PHIEUMUON.MASV = '" + username.ToString() + "'");
            if (tbThongBao.Rows.Count > 0)
            {
                for (int i = 0; i < tbThongBao.Rows.Count; i++)
                {
                    int ngay; string thang; string nam;
                    ngay = int.Parse(tbThongBao.Rows[i]["NGAYHENTRA"].ToString().Substring(0, 2)) + 1;
                    thang = tbThongBao.Rows[i]["NGAYHENTRA"].ToString().Substring(3, 2);
                    nam = tbThongBao.Rows[i]["NGAYHENTRA"].ToString().Substring(6, 5);

                    thongbao = "*Thông báo ngày: " + ngay.ToString() + " / " + thang + " / " + nam + "\n  Sách: " + tbThongBao.Rows[i]["TENSACH"].ToString() + "\n  -Ngày mượn: " + tbThongBao.Rows[i]["NGAYMUON"].ToString() + "\n  -Ngày hẹn trả: " + tbThongBao.Rows[i]["NGAYHENTRA"].ToString() + "\n";

                }
                thongbao = thongbao + "Đã quá hạn trả sách. Vui lòng mang sách đến thư viện để trả sách sớm \nnhất có thể!\nNếu quá hạn 5 tháng, tài khoản của bạn sẽ bị khóa!";
            }
            lblThongBao.Text = thongbao;

            ////
            string thongbao2 = "";
            DataTable tbThongBao2 = dtBase.DataSelect("select * from PHIEUMUON join SACH on PHIEUMUON.MASACH = SACH.MASACH where PHIEUMUON.TRANGTHAI like N'%Không trả sách%' and PHIEUMUON.MASV = '" + username.ToString() + "'");
            if (tbThongBao2.Rows.Count > 0)
            {
                int ngay0; string thang0; string nam0;
                ngay0 = int.Parse(tbThongBao2.Rows[0]["NGAYHENTRA"].ToString().Substring(0, 2)) + 1;
                thang0 = tbThongBao2.Rows[0]["NGAYHENTRA"].ToString().Substring(3, 2);
                nam0 = tbThongBao2.Rows[0]["NGAYHENTRA"].ToString().Substring(6, 5);

                int thangtg = 0, namtg = 0;
                if((int.Parse(thang0) + 5) > 12)
                {
                    thangtg = int.Parse(thang0) + 5 - 12;
                    namtg = int.Parse(nam0) + 1;
                }
                else if((int.Parse(thang0) + 5) <= 12)
                {
                    thangtg = int.Parse(thang0) + 5;
                    namtg = int.Parse(nam0);
                }
                thongbao2 = "*Thông báo ngày: " + ngay0.ToString() + " / " + thangtg.ToString() + " / " + namtg.ToString() + "\n  Sách: " + tbThongBao2.Rows[0]["TENSACH"].ToString() + "\n  -Ngày mượn: " + tbThongBao2.Rows[0]["NGAYMUON"].ToString() + "\n  -Ngày hẹn trả: " + tbThongBao2.Rows[0]["NGAYHENTRA"].ToString() + "\n";

                thongbao2 = thongbao2 + "Tài khoản của bạn đã bị khóa!\nLý do: Đã quá hạn trả sách là 5 tháng.!\nNếu bạn muốn tiếp tục mượn sách, vui lòng thanh toán phí nộp phạt\n qua ứng dụng ViettelPay và đến thư viện nhờ nhân viên mở tài khoản!";
            }
            lblKtra.Text = thongbao2;

            ////
            string thongbaoMoTK = "";
            DataTable tbThongBaoMoTK = dtBase.DataSelect("select * from TAIKHOAN_MO where MASV_MO = '" + username.ToString() + "'");
            if (tbThongBaoMoTK.Rows.Count > 0)
            {
                int ngay0; int thang0; int nam0;
                ngay0 = int.Parse(tbThongBaoMoTK.Rows[0]["NGAYMO_TK"].ToString().Substring(0, 2));
                thang0 = int.Parse(tbThongBaoMoTK.Rows[0]["NGAYMO_TK"].ToString().Substring(3, 2));
                nam0 = int.Parse(tbThongBaoMoTK.Rows[0]["NGAYMO_TK"].ToString().Substring(6, 5));

                thongbaoMoTK = "*Thông báo ngày: " + ngay0.ToString() + " / " + thang0.ToString() + " / " + nam0.ToString() + "\n";

                thongbaoMoTK = thongbaoMoTK + "Tài khoản của bạn đã được mở! Bạn đã có thể tiếp tục mượn sách!";
            }
            lblMoTK.Text = thongbaoMoTK;


            //////
            lblQuyDinh.Text = "-Bạn có thể mượn tối đa 3 quyển sách khác loại. Mỗi loại sách bạn chỉ được mượn 1 quyển.\n" +
                "-Vi phạm: +Không vi phạm(Trả sách đúng hạn, không làm hỏng hoặc mất sách): Không bị phạt.\n" +
                "          +Trả đúng hạn, Làm hỏng sách: Nộp phạt 20% giá bìa sách.\n" +
                "          +Trả đúng hạn, Làm mất sách: Nộp phạt 50% giá bìa sách.\n" +
                "          +Trả quá hạn: Trừ điểm rèn luyện.\n" +
                "          +Trả quá hạn, Làm hỏng sách: Nộp phạt 25% giá bìa sách\n" +
                "          +Trả quá hạn, Làm mất sách: Nộp phạt 55% giá bìa sách\n" +
                "          +Không trả sách: Khóa tài khoản và nộp phạt 70% giá bìa sách.";

        }
        private void frmSinhVien_Load(object sender, EventArgs e)
        {
            ///sinh viên
            DataTable tbSV = dtBase.DataSelect("select * from SINHVIEN where MASV = '" + username.ToString() + "'");
            lblXinChao.Text = lblXinChao.Text + tbSV.Rows[0]["HOTEN"].ToString();

            txtMaSV.Text = tbSV.Rows[0]["MASV"].ToString();
            txtPass.Text = tbSV.Rows[0]["PASS"].ToString();
            txtHoTen.Text = tbSV.Rows[0]["HOTEN"].ToString();
            txtNgaySinh.Text = tbSV.Rows[0]["NGAYSINH"].ToString().Substring(0, 11);
            cmbGioiTinh.Text = tbSV.Rows[0]["GIOITINH"].ToString();
            txtDiaChi.Text = tbSV.Rows[0]["DIAVHI"].ToString();
            txtSDT.Text = tbSV.Rows[0]["SDT"].ToString();
            txtEmail.Text = tbSV.Rows[0]["EMAIL"].ToString();

            ///sách
            LoadDataSach();
            dgvSach.Columns[0].HeaderText = "Mã Sách";
            dgvSach.Columns[1].HeaderText = "Tên Sách";
            dgvSach.Columns[2].HeaderText = "Tác giả";
            dgvSach.Columns[3].HeaderText = "NXB";
            dgvSach.Columns[4].HeaderText = "Thể loại";
            dgvSach.Columns[5].HeaderText = "Số lượng";
            dgvSach.Columns[6].HeaderText = "Giá bìa";


            dgvSach.Columns[1].Width = 300;
            dgvSach.Columns[2].Width = 200;
            dgvSach.Columns[3].Width = 200;
            dgvSach.Columns[4].Width = 200;


            ///phiếu mượn
            LoadDataPhieuMuon();
            dgvMuon.Columns[0].HeaderText = "Mã phiếu mượn";
            dgvMuon.Columns[1].HeaderText = "Tên sách";
            dgvMuon.Columns[2].HeaderText = "Số lượng mượn";
            dgvMuon.Columns[3].HeaderText = "Ngày mượn";
            dgvMuon.Columns[4].HeaderText = "Ngày hẹn trả";
            dgvMuon.Columns[5].HeaderText = "Trạng thái mượn";

            dgvMuon.Columns[1].Width = 300;
            dgvMuon.Columns[5].Width = 300;


            ///phiếu trả
            LoadDataPhieuTra();
            dgvPhieuTra.Columns[0].HeaderText = "Mã phiếu trả";
            dgvPhieuTra.Columns[1].HeaderText = "Mã phiếu mượn";
            dgvPhieuTra.Columns[2].HeaderText = "Tên sách";
            dgvPhieuTra.Columns[3].HeaderText = "Ngày trả sách";
            dgvPhieuTra.Columns[4].HeaderText = "Nội dung vi phạm";
            dgvPhieuTra.Columns[5].HeaderText = "Nội dung phạt";
            dgvPhieuTra.Columns[6].HeaderText = "Phí nộp phạt";
            dgvPhieuTra.Columns[7].HeaderText = "Ghi chú";


            dgvPhieuTra.Columns[2].Width = 300;
            dgvPhieuTra.Columns[4].Width = 200;
            dgvPhieuTra.Columns[5].Width = 250;
            dgvPhieuTra.Columns[7].Width = 250;


            loadDataThongBao();
            btnLuu.Enabled = false;
        }

        ///Thông tin cá nhân sv
        private void txtSDT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Convert.ToInt16(e.KeyChar) < Convert.ToInt16('0') || Convert.ToInt16(e.KeyChar) > Convert.ToInt16('9')) &&
                Convert.ToInt16(e.KeyChar) != 8)
            {
                MessageBox.Show("Bạn chỉ được nhập số nguyên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Handled = true; 
            }
        }

        private void btnSuaTT_Click(object sender, EventArgs e)
        {
            txtHoTen.Enabled = true;
            txtPass.Enabled = true;
            txtDiaChi.Enabled = true;
            txtSDT.Enabled = true;
            txtEmail.Enabled = true;
            btnLuu.Enabled = true;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (txtHoTen.Text.Trim() == "" || txtPass.Text.Trim() == "" || txtDiaChi.Text.Trim() == "" || txtSDT.Text.Trim() == "" || txtEmail.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đủ thông tin để tiến hành sửa! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPass.Focus();
                return;
            }
            if (txtPass.Text.Length < 8)
            {
                MessageBox.Show("Mật khẩu cần có từ 8 ký tự trở lên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            if (txtEmail.ToString().Contains("@gmail.com") != true && txtEmail.ToString().Contains("@st.utc.edu.vn") != true)
            {
                MessageBox.Show("Email không đúng định dạng!\nVui lòng nhập đúng email!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtEmail.Focus();
                return;
            }
            else
            {
                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    
                    string sqlUpdatePass = "update SINHVIEN set PASS=(N'" + txtPass.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdatePass);

                    string sqlUpdateHT = "update SINHVIEN set HOTEN=(N'" + txtHoTen.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdateHT);

                    string sqlUpdateGTinh = "update SINHVIEN set GIOITINH=(N'" + cmbGioiTinh.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdateGTinh);

                    string sqlUpdateSDT = "update SINHVIEN set SDT=(N'" + txtSDT.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdateSDT);

                    string sqlUpdateDC = "update SINHVIEN set DIAVHI=(N'" + txtDiaChi.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdateDC);

                    string sqlUpdateEmail = "update SINHVIEN set EMAIL=(N'" + txtEmail.Text + "') where MASV = (N'" + username.ToString() + "')";
                    dtBase.Updatedate(sqlUpdateEmail);

                    txtHoTen.Enabled = false;
                    txtPass.Enabled = false;
                    txtDiaChi.Enabled = false;
                    txtSDT.Enabled = false;
                    txtEmail.Enabled = false;
                    btnLuu.Enabled = false;
                }     
            }
        }

        ///sách
        private void txtTKSach_TextChanged(object sender, EventArgs e)
        {
            dgvSach.DataSource = dtBase.DataSelect("Select * from SACH where MASACH like (N'%" + txtTKSach.Text + "%') or TENSACH like (N'%" + txtTKSach.Text + "%')");
        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            LoadDataSach();
            txtTKSach.Text = "";
        }
    }
}
