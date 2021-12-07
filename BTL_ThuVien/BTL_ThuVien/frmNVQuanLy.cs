using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BTL_ThuVien
{
    public partial class frmNVQuanLy : Form
    {
        Module.DataAccess dtBase = new Module.DataAccess();
        Module.Function sinhma = new Module.Function();
        DateTime now = DateTime.Today;

        public frmNVQuanLy()
        {
            InitializeComponent();
        }
        void LoadDataSV()
        {
            dgvSinhVien.DataSource = dtBase.DataSelect("select MASV, HOTEN, NGAYSINH, GIOITINH, DIAVHI, SDT, EMAIL from SINHVIEN");
        }

        void LoadDataSach()
        {
            dgvSach.DataSource = dtBase.DataSelect("Select SACH.MASACH, SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from ((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB");
        }

        void LoadDataTG()
        {
            dgvTacGia.DataSource = dtBase.DataSelect("Select * from TACGIA");
        }
        void LoadDataTheLoai()
        {
            dgvTheLoai.DataSource = dtBase.DataSelect("Select * from THELOAI");
        }
        void LoadDataNXB()
        {
            dgvNXB.DataSource = dtBase.DataSelect("Select * from NXB");
        }

        void LoadDataPhieuMuon()
        {
            dgvPhieuMuon.DataSource = dtBase.DataSelect("Select PHIEUMUON.MAPHIEUMUON, PHIEUMUON.MASV, SinhVien.HOTEN, SACH.TENSACH, " +
                "PHIEUMUON.SOLUONGMUON, PHIEUMUON.NGAYMUON, PHIEUMUON.NGAYHENTRA, PHIEUMUON.TRANGTHAI " +
                " from (PHIEUMUON left join SinhVien on PHIEUMUON.MASV = SinhVien.MASV) left join SACH on PHIEUMUON.MASACH = SACH.MASACH" +
                " where month(PHIEUMUON.NGAYMUON)= '" + now.Month.ToString() + "' " +
                "and  year(PHIEUMUON.NGAYMUON)= '" + now.Year.ToString() + "'");
        }

        void LoadDataPhieuTra()
        {
            dgvPhieuTra.DataSource = dtBase.DataSelect("Select PHIEUTRA.MAPHIEUTRA, PHIEUTRA.MAPHIEUMUON, SinhVien.MASV, SinhVien.HOTEN, SACH.TENSACH, PHIEUTRA.NGAYTRA, VIPHAM.ND_VIPHAM, VIPHAM.ND_PHAT, PHINOPPHAT, GHICHU " +
                "from ((((PHIEUTRA LEFT JOIN VIPHAM ON PHIEUTRA.MAVIPHAM = VIPHAM.MAVIPHAM) join PHIEUMUON on PHIEUMUON.MAPHIEUMUON = PHIEUTRA.MAPHIEUMUON) join SinhVien on PHIEUMUON.MASV = SinhVien.MaSV) join Sach on SACH.MASACH = PHIEUMUON.MASACH) " +
                "where month(PHIEUTRA.NGAYTRA)= '" + now.Month.ToString() + "' " +
                "and  year(PHIEUTRA.NGAYTRA)= '" + now.Year.ToString() + "'");
        }

        void TinhTongPM()
        {
            DataTable PM = dtBase.DataSelect("Select * from PHIEUMUON");
            if (PM.Rows.Count > 0)
            {
                lblTongPM.Text = PM.Rows.Count.ToString();
            }
            else
            {
                lblTongPM.Text = "0";
            }

            DataTable PM_DTra = dtBase.DataSelect("Select * from PHIEUMUON where TRANGTHAI = N'Đã trả'");
            if (PM_DTra.Rows.Count > 0)
            {
                lblPM_DTra.Text = PM_DTra.Rows.Count.ToString();
            }
            else
            {
                lblPM_DTra.Text = "0";
            }

            DataTable PM_DM = dtBase.DataSelect("Select * from PHIEUMUON where TRANGTHAI = N'Đang mượn sách'");
            if (PM_DM.Rows.Count > 0)
            {
                lblPM_DM.Text = PM_DM.Rows.Count.ToString();
            }
            else
            {
                lblPM_DM.Text = "0";
            }

            DataTable PM_QH = dtBase.DataSelect("Select * from PHIEUMUON where TRANGTHAI = N'Quá hạn'");
            if (PM_QH.Rows.Count > 0)
            {
                lblPM_QH.Text = PM_QH.Rows.Count.ToString();
            }
            else
            {
                lblPM_QH.Text = "0";
            }

            DataTable PM_KTra = dtBase.DataSelect("Select * from PHIEUMUON where TRANGTHAI = N'Không trả sách'");
            if (PM_KTra.Rows.Count > 0)
            {
                lblPM_KTra.Text = PM_KTra.Rows.Count.ToString();
            }
            else
            {
                lblPM_KTra.Text = "0";
            }

            DataTable PMSV = dtBase.DataSelect("select DISTINCT(MASV) from PHIEUMUON");
            if (PMSV.Rows.Count > 0)
            {
                lblTongSVMuon.Text = PMSV.Rows.Count.ToString();
            }
            else
            {
                lblTongSVMuon.Text = "0";
            }
        }
        void TinhTongPT()
        {
            int TongPT = 0;
            int TongPTKVP = 0;
            DataTable PT = dtBase.DataSelect("select * from PHIEUTRA");
            if (PT.Rows.Count > 0)
            {
                lblTongPT.Text = PT.Rows.Count.ToString();
                TongPT = int.Parse(PT.Rows.Count.ToString());
            }
            else
            {
                lblTongPT.Text = "0";
            }

            DataTable PTKVP = dtBase.DataSelect("select * from PHIEUTRA where MAVIPHAM = 'VP0'");
            if (PTKVP.Rows.Count > 0)
            {
                lblK_ViPham.Text = PTKVP.Rows.Count.ToString();
                TongPTKVP = int.Parse(PTKVP.Rows.Count.ToString());
            }
            else
            {
                lblK_ViPham.Text = "0"; 
            }

            lblC_ViPham.Text = (TongPT - TongPTKVP).ToString();
        }
        void TinhTongSach()
        {
            string strLocSach = "Select DISTINCT(SACH.MASACH), SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from (((((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB) left join PHIEUMUON on PHIEUMUON.MASACH = SACH.MASACH) " +
                "LEFT join PHIEUTRA on PHIEUTRA.MAPHIEUMUON = PHIEUMUON.MAPHIEUMUON) LEFT JOIN VIPHAM on VIPHAM.MAVIPHAM = PHIEUTRA.MAVIPHAM";

            string strLocSachM = "Select DISTINCT(SACH.MASACH), SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
               "from ((((PHIEUMUON LEFT JOIN SACH ON PHIEUMUON.MASACH = SACH.MASACH) left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) join NXB ON SACH.MA_NXB = NXB.MA_NXB) ";

            DataTable TongSach = dtBase.DataSelect("Select SACH.MASACH, SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from ((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB");

            DataTable SachMuon = dtBase.DataSelect(strLocSachM + " where PHIEUMUON.TRANGTHAI = N'Đang mượn sách' or PHIEUMUON.TRANGTHAI = N'Quá hạn' or PHIEUMUON.TRANGTHAI = N'Không trả sách'"); 

            DataTable SMQuaHan = dtBase.DataSelect(strLocSachM + " where PHIEUMUON.TRANGTHAI = N'Quá hạn' or PHIEUMUON.TRANGTHAI = N'Không trả sách'");

            DataTable SachHong = dtBase.DataSelect(strLocSach + " where VIPHAM.ND_VIPHAM like N'%Làm hỏng sách%'");

            if(TongSach.Rows.Count > 0)
            {
                lbTongSach.Text = TongSach.Rows.Count.ToString();
            }
            else
            {
                lbTongSach.Text = "0";
            }

            if(SachMuon.Rows.Count > 0)
            {
                lblTongSDM.Text = SachMuon.Rows.Count.ToString();
            }
            else
            {
                lblTongSDM.Text = "";
            }

            if(SMQuaHan.Rows.Count > 0)
            {
                lblSachQH.Text = SMQuaHan.Rows.Count.ToString();
            }
            else
            {
                lblSachQH.Text = "0";
            }

            if(SachHong.Rows.Count > 0)
            {
                lblTongSachHong.Text = SachHong.Rows.Count.ToString();
            }
            else
            {
                lblTongSachHong.Text = "0";
            }
        }
        void loadCmbTacGia() //sách
        {
            cmbTacGia.DataSource = dtBase.DataSelect("select MATACGIA, TENTG from TACGIA");
            cmbTacGia.DisplayMember = "TENTG";
            cmbTacGia.ValueMember = "MATACGIA";
        }
        void loadCmbTheLoai() ///sách
        {
            cmbTheLoai.DataSource = dtBase.DataSelect("select MATHELOAI, TENTHELOAI from THELOAI");
            cmbTheLoai.DisplayMember = "TENTHELOAI";
            cmbTheLoai.ValueMember = "MATHELOAI";
        }
        void loadCmbNXB() //sách
        {
            cmbNXB.DataSource = dtBase.DataSelect("select MA_NXB, TENNXB from NXB");
            cmbNXB.DisplayMember = "TENNXB";
            cmbNXB.ValueMember = "MA_NXB";
        }
        void loadCmbSach() ///phiếu mượn
        {
            cmbSach.DataSource = dtBase.DataSelect("select MASACH, TENSACH from SACH");
            cmbSach.DisplayMember = "TENSACH";
            cmbSach.ValueMember = "MASACH";
        }
        private void frmNVQuanLy_Load(object sender, EventArgs e)
        {

            String dataPM = "update PHIEUMUON set TRANGTHAI = N'Quá hạn' where TRANGTHAI = N'Đang mượn sách' and NGAYHENTRA < '" + now.ToString("yyyy-MM-dd") + "'";
            dtBase.Updatedate(dataPM);
            LoadDataPhieuMuon();

            String dataPM_K = "update PHIEUMUON set TRANGTHAI = N'Không trả sách' where TRANGTHAI = N'Quá hạn' and DATEDIFF(day, NGAYHENTRA, '" + now.ToString("yyyy-MM-dd") + "') >= 150"; //quá hạn 5 tháng
            dtBase.Updatedate(dataPM_K);
            LoadDataPhieuMuon();

            ////Sinh Viên
            LoadDataSV();
            dgvSinhVien.Columns[0].HeaderText = "Mã Sinh Viên";
            dgvSinhVien.Columns[1].HeaderText = "Họ Tên";
            dgvSinhVien.Columns[2].HeaderText = "Ngày Sinh";
            dgvSinhVien.Columns[3].HeaderText = "Giới Tính";
            dgvSinhVien.Columns[4].HeaderText = "Địa chỉ";
            dgvSinhVien.Columns[5].HeaderText = "SĐT";
            dgvSinhVien.Columns[6].HeaderText = "Email";

            dgvSinhVien.Columns[1].Width = 200;
            dgvSinhVien.Columns[4].Width = 200;
            dgvSinhVien.Columns[6].Width = 200;

            ///////////SÁCH

            TinhTongSach();
            loadCmbTacGia();
            loadCmbTheLoai();
            loadCmbNXB();

            LoadDataSach();

            cmbTacGia.SelectedIndex = -1;
            cmbTheLoai.SelectedIndex = -1;
            cmbNXB.SelectedIndex = -1;
            dgvSach.Columns[0].HeaderText = "Mã Sách";
            dgvSach.Columns[1].HeaderText = "Tên Sách";
            dgvSach.Columns[2].HeaderText = "Tác giả";
            dgvSach.Columns[3].HeaderText = "NXB";
            dgvSach.Columns[4].HeaderText = "Thể Loại";
            dgvSach.Columns[5].HeaderText = "Số lượng";
            dgvSach.Columns[6].HeaderText = "Giá bìa";


            dgvSach.Columns[1].Width = 300;
            dgvSach.Columns[2].Width = 200;
            dgvSach.Columns[3].Width = 200;
            dgvSach.Columns[4].Width = 200;


            LoadDataTG();
            dgvTacGia.Columns[0].HeaderText = "Mã Tác Giả";
            dgvTacGia.Columns[1].HeaderText = "Tên Tác Giả";


            LoadDataTheLoai();
            dgvTheLoai.Columns[0].HeaderText = "Mã thể loại";
            dgvTheLoai.Columns[1].HeaderText = "Tên thể loại";

            LoadDataNXB();
            dgvNXB.Columns[0].HeaderText = "Mã NXB";
            dgvNXB.Columns[1].HeaderText = "Tên NXB";
            dgvNXB.Columns[2].HeaderText = "Địa chỉ NXB";
            dgvNXB.Columns[3].HeaderText = "SĐT NXB";

            dgvNXB.Columns[1].Width = 200;
            dgvNXB.Columns[2].Width = 250;

            ////phieu muon
            loadCmbSach();

            cmbThang.Text = now.Month.ToString();
            cmbNam.Text = now.Year.ToString();

            TinhTongPM();

            LoadDataPhieuMuon();

            cmbSach.SelectedIndex = -1;
            dgvPhieuMuon.Columns[0].HeaderText = "Mã phiếu mượn";
            dgvPhieuMuon.Columns[1].HeaderText = "Mã sinh viên";
            dgvPhieuMuon.Columns[2].HeaderText = "Tên Sinh Viên";
            dgvPhieuMuon.Columns[3].HeaderText = "Sách";
            dgvPhieuMuon.Columns[4].HeaderText = "Số lượng";
            dgvPhieuMuon.Columns[5].HeaderText = "Ngày mượn";
            dgvPhieuMuon.Columns[6].HeaderText = "Ngày hẹn trả";
            dgvPhieuMuon.Columns[7].HeaderText = "Trạng thái mượn";

            dgvPhieuMuon.Columns[2].Width = 200;
            dgvPhieuMuon.Columns[3].Width = 300;
            dgvPhieuMuon.Columns[7].Width = 300;


            ///phiếu trả
            ///
            cmbThangPT.Text = now.Month.ToString();
            cmbNamPT.Text = now.Year.ToString();

            TinhTongPT();

            //cmbLocPT.DataSource = dtBase.DataSelect("select ND_VIPHAM from VIPHAM");
            //cmbLocPT.DisplayMember = "ND_VIPHAM";
            //cmbLocPT.ValueMember = "MAVIPHAM";
            LoadDataPhieuTra();
            cmbViPham.SelectedIndex = -1;
            dgvPhieuTra.Columns[0].HeaderText = "Mã phiếu trả";
            dgvPhieuTra.Columns[1].HeaderText = "Mã phiếu mượn";
            dgvPhieuTra.Columns[2].HeaderText = "Mã Sinh Viên";
            dgvPhieuTra.Columns[3].HeaderText = "Tên Sinh viên";
            dgvPhieuTra.Columns[4].HeaderText = "Tên Sách";
            dgvPhieuTra.Columns[5].HeaderText = "Ngày trả sách";
            dgvPhieuTra.Columns[6].HeaderText = "Nội dung vi phạm";
            dgvPhieuTra.Columns[7].HeaderText = "Nội dung phạt";
            dgvPhieuTra.Columns[8].HeaderText = "Phí nộp phạt";
            dgvPhieuTra.Columns[9].HeaderText = "Ghi chú";

            dgvPhieuTra.Columns[3].Width = 200;
            dgvPhieuTra.Columns[4].Width = 300;
            dgvPhieuTra.Columns[6].Width = 200;
            dgvPhieuTra.Columns[7].Width = 250;
            dgvPhieuTra.Columns[9].Width = 200;

            btnLuuSach.Enabled = false;
            btnUpSach.Enabled = false;
            btnLuuSach.BackColor = Color.PeachPuff;
            btnUpSach.BackColor = Color.PeachPuff;

            btnLuuTG.Enabled = false;
            btnUpTG.Enabled = false;
            btnLuuTG.BackColor = Color.PeachPuff;
            btnUpTG.BackColor = Color.PeachPuff;

            btnLuuTL.Enabled = false;
            btnUpTL.Enabled = false;
            btnLuuTL.BackColor = Color.PeachPuff;
            btnUpTL.BackColor = Color.PeachPuff;

            btnLuuNXB.Enabled = false;
            btnUpNXB.Enabled = false;
            btnLuuNXB.BackColor = Color.PeachPuff;
            btnUpNXB.BackColor = Color.PeachPuff;

            btnLuuPM.Enabled = false;
            btnUpPM.Enabled = false;
            btnLuuPM.BackColor = Color.PeachPuff;
            btnUpPM.BackColor = Color.PeachPuff;

            btnLuuPT.Enabled = false;
            btnUpPT.Enabled = false;
            btnLuuPT.BackColor = Color.PeachPuff;
            btnUpPT.BackColor = Color.PeachPuff;


        }
        private void btnThoatSV_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                this.Close();
            }
        }

        ///Sinh viên
        public void ResetValues_SV()
        {
            txtMaSV_SV.Text = "";
            txtHoTen.Text = "";
            txtSDT.Text = "";
            txtDiaChi.Text = "";
            cmbGioiTinh.SelectedIndex = -1;
            txtEmail.Text = "";
            dtpNgaySinh.Text = now.ToString();
            cmbLocSV.Text = "";
            txtMSV_MoTK.Text = "";
            txtTenSVMoTK.Text = "";
        }
        private void btnReloadSV_Click(object sender, EventArgs e)
        {
            txtTKSV.Text = "";
            ResetValues_SV();
            LoadDataSV();
        }
        private void cmbLocSV_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cmbLocSV.Text == "Tất cả")
            {
                dgvSinhVien.DataSource = dtBase.DataSelect("select MASV, HOTEN, NGAYSINH, GIOITINH, DIAVHI, SDT, EMAIL from SINHVIEN");
            }
            else if(cmbLocSV.Text == "Sinh viên bị khóa tài khoản")
            {
                dgvSinhVien.DataSource = dtBase.DataSelect("select MASV, HOTEN, NGAYSINH, GIOITINH, DIAVHI, SDT, EMAIL " +
                    "from SINHVIEN join TAIKHOAN_BIKHOA on SINHVIEN.MASV = TAIKHOAN_BIKHOA.MASV_BIKHOA");
            }
            else if (cmbLocSV.Text == "Sinh viên đã mở tài khoản")
            {
                dgvSinhVien.DataSource = dtBase.DataSelect("select MASV, HOTEN, NGAYSINH, GIOITINH, DIAVHI, SDT, EMAIL " +
                    "from SINHVIEN join TAIKHOAN_MO on SINHVIEN.MASV = TAIKHOAN_MO.MASV_MO");
            }
        }
        private void btnXoaSV_Click(object sender, EventArgs e)
        {
            if (txtMaSV_SV.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã sinh viên!\nVui lòng nhập mã sinh viên! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaSV_SV.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSachSV = dtBase.DataSelect("select * from SINHVIEN where MASV='" + txtMaSV_SV.Text + "'");
                if (tbDanhSachSV.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaSV_SV.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSV_SV.Focus();
                    ResetValues_SV();
                    return;
                }
                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDelete = "delete from SINHVIEN where MASV = (N'" + txtMaSV_SV.Text + "')";
                    dtBase.Updatedate(sqlDelete);
                    LoadDataSV();
                    ResetValues_SV();
                }
            }
        }
        private void txtTKSV_TextChanged(object sender, EventArgs e)
        {
            dgvSinhVien.DataSource = dtBase.DataSelect("Select MASV, HOTEN, NGAYSINH, GIOITINH, DIAVHI, SDT, EMAIL from SINHVIEN where MASV like (N'%" + txtTKSV.Text + "%') or HOTEN like (N'%" + txtTKSV.Text + "%')");
        }
        private void txtMSV_MoTK_TextChanged(object sender, EventArgs e)
        {
            DataTable tbTKMo = dtBase.DataSelect("select * from SINHVIEN where MASV = '" + txtMSV_MoTK.Text + "'");
            if (tbTKMo.Rows.Count > 0)
            {
                txtTenSVMoTK.Text = tbTKMo.Rows[0]["HOTEN"].ToString();
            }
            else
            {
                txtTenSVMoTK.Text = "";
            }
        }
        private void btnMTKSV_Click(object sender, EventArgs e)
        {
            if(txtMSV_MoTK.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã sinh viên!\nVui lòng nhập mã sinh viên mà bạn muốn mở tài khoản!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMSV_MoTK.Focus();
                return;
            }
            DataTable tbTK_Mo = dtBase.DataSelect("select * from TAIKHOAN_MO where MASV_MO = '" + txtMSV_MoTK.Text + "'");
            if (tbTK_Mo.Rows.Count > 0)
            {
                MessageBox.Show("Bạn không thể mở tài khoản này!\nSinh viên chỉ được mở lại tài khoản một lần duy nhất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtMSV_MoTK.Focus();
                return;
            }

            DataTable tkSV = dtBase.DataSelect("select * from TAIKHOAN_BIKHOA where MASV_BIKHOA = '"+ txtMSV_MoTK.Text + "'");
            if(tkSV.Rows.Count == 0)
            {
                MessageBox.Show("Tài khoản này không đúng hoặc không bị khóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMSV_MoTK.Focus();
                return;
            }

            if (dtpNgayMoTK.Value.ToString().Substring(0, 11) != now.ToString().Substring(0, 11))
            {
                MessageBox.Show("Ngày mở tài khoản phải là ngày hôm nay!\nBạn không được thay đổi ngày mở tài khoản này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dtpNgayMoTK.Focus();
                return;
            }

            DataTable ktraPhiNopPhat = dtBase.DataSelect("select * from PHIEUTRA join PHIEUMUON on PHIEUTRA.MAPHIEUMUON = PHIEUMUON.MAPHIEUMUON where PHIEUMUON.MASV = '" + txtMSV_MoTK.Text + "' and GHICHU = N'Chưa thanh toán phí nộp phạt'");
            if(ktraPhiNopPhat.Rows.Count > 0)
            {
                MessageBox.Show("Bạn chưa hoàn thành phí nộp phạt!\nBạn không thể mở tài khoản này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMSV_MoTK.Focus();
                return;
            }

            if (tkSV.Rows.Count > 0 && tbTK_Mo.Rows.Count == 0)
            {
                string deleteTK = "delete TAIKHOAN_BIKHOA where MASV_BIKHOA = '" + txtMSV_MoTK.Text + "'";
                dtBase.Updatedate(deleteTK);

                string insertTK = "insert into TAIKHOAN_MO values(N'" + txtMSV_MoTK.Text + "', N'" + dtpNgayMoTK.Value.ToString("yyyy-MM-dd") + "')";
                dtBase.Updatedate(insertTK);
                MessageBox.Show("Bạn đã mở tài khoản thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMSV_MoTK.Text = "";
                txtTenSVMoTK.Text = "";
                LoadDataSV();
            }
        }

        /// Sach
        public void ResetValues_Sach()
        {
            txtMaSach.Text = "";
            txtTenSach.Text = "";
            cmbTacGia.SelectedIndex = -1;
            cmbTheLoai.SelectedIndex = -1;
            cmbNXB.SelectedIndex = -1;
            txtSLSach.Text = "";
            txtGiaBia.Text = "";
            cmbLocSah.Text = "";
        }
        private void txtSLSach_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Convert.ToInt16(e.KeyChar) < Convert.ToInt16('0') || Convert.ToInt16(e.KeyChar) > Convert.ToInt16('9'))
                && Convert.ToInt16(e.KeyChar) != 8)
            {
                MessageBox.Show("Bạn chỉ được nhập số", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Handled = true;
            }
        }

        private void txtGiaBia_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Convert.ToInt16(e.KeyChar) < Convert.ToInt16('0') || Convert.ToInt16(e.KeyChar) > Convert.ToInt16('9'))
                && Convert.ToInt16(e.KeyChar) != 8 && Convert.ToInt16(e.KeyChar) != Convert.ToInt16('.'))
            {
                MessageBox.Show("Bạn chỉ được nhập số", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Handled = true;
            }
        }

        private void cmbLocSah_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strLocSach = "Select DISTINCT(SACH.MASACH), SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from (((((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB) left join PHIEUMUON on PHIEUMUON.MASACH = SACH.MASACH) " +
                "LEFT join PHIEUTRA on PHIEUTRA.MAPHIEUMUON = PHIEUMUON.MAPHIEUMUON) LEFT JOIN VIPHAM on VIPHAM.MAVIPHAM = PHIEUTRA.MAVIPHAM";

            string strLocSachM = "Select DISTINCT(SACH.MASACH), SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
               "from ((((PHIEUMUON LEFT JOIN SACH ON PHIEUMUON.MASACH = SACH.MASACH) left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) join NXB ON SACH.MA_NXB = NXB.MA_NXB) ";

            if (cmbLocSah.Text == "Tất cả")
            {
                dgvSach.DataSource = dtBase.DataSelect("Select SACH.MASACH, SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from ((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB");
            }
            else if(cmbLocSah.Text == "Sách bị hỏng")
            {
                dgvSach.DataSource = dtBase.DataSelect(strLocSach + " where VIPHAM.ND_VIPHAM like N'%Làm hỏng sách%'");
            }
            else if(cmbLocSah.Text == "Sách đã và đang được mượn")
            {
                dgvSach.DataSource = dtBase.DataSelect(strLocSachM);
            }
            else if (cmbLocSah.Text == "Sách đang được mượn")
            {
                dgvSach.DataSource = dtBase.DataSelect(strLocSachM + " where PHIEUMUON.TRANGTHAI = N'Đang mượn sách' or  PHIEUMUON.TRANGTHAI = N'Quá hạn' or  PHIEUMUON.TRANGTHAI = N'Không trả sách'");
            }
            else if (cmbLocSah.Text == "Sách đang được mượn nhưng quá hạn trả")
            {
                dgvSach.DataSource = dtBase.DataSelect(strLocSachM + " where PHIEUMUON.TRANGTHAI = N'Quá hạn' or  PHIEUMUON.TRANGTHAI = N'Không trả sách'");
            }
        }

        void reloadSach()
        {
            txtMaSach.Text = "";
            txtMaSach.Enabled = true;
            btnThemSach.Enabled = true;
            btnSuaSach.Enabled = true;
            btnXoaSach.Enabled = true;
            btnLuuSach.Enabled = false;
            btnUpSach.Enabled = false;

            btnThemSach.BackColor = Color.Peru;
            btnSuaSach.BackColor = Color.Peru;
            btnXoaSach.BackColor = Color.Peru;
            btnLuuSach.BackColor = Color.PeachPuff;
            btnUpSach.BackColor = Color.PeachPuff;

        }
        private void btnReloadSach_Click(object sender, EventArgs e)
        {
            txtTKSach.Text = "";
            ResetValues_Sach();
            LoadDataSach();
            reloadSach();
        }
        private void btnThemSach_Click(object sender, EventArgs e)
        {
            txtMaSach.Text = sinhma.SinhMa("SACH", "MS", "MASACH");
            txtMaSach.Enabled = false;

            txtTenSach.Text = "";
            cmbTacGia.SelectedIndex = -1;
            cmbTheLoai.SelectedIndex = -1;
            cmbNXB.SelectedIndex = -1;
            txtSLSach.Text = "";
            txtGiaBia.Text = "";

            btnSuaSach.Enabled = false;
            btnXoaSach.Enabled = false;
            btnLuuSach.Enabled = true;

            btnSuaSach.BackColor = Color.PeachPuff;
            btnXoaSach.BackColor = Color.PeachPuff;
            btnLuuSach.BackColor = Color.Peru;

        }
        private void btnLuuSach_Click(object sender, EventArgs e)
        {
            if (txtTenSach.Text.Trim() == "" || txtSLSach.Text.Trim() == "" || cmbTacGia.SelectedIndex == -1 || cmbTheLoai.SelectedIndex == -1 || cmbNXB.SelectedIndex == -1 || txtGiaBia.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin để tiến hành thêm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                string sqlInsert = "insert into SACH values(N'" + txtMaSach.Text + "', N'" + txtTenSach.Text + "', N'" + cmbTacGia.SelectedValue.ToString() + "', N'" + cmbNXB.SelectedValue.ToString() + "', N'" + cmbTheLoai.SelectedValue.ToString() + "', '" + int.Parse(txtSLSach.Text) + "', '" + float.Parse(txtGiaBia.Text) + "')";
                dtBase.Updatedate(sqlInsert);
                MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDataSach();
                ResetValues_Sach();
                loadCmbSach();
                cmbSach.SelectedIndex = -1;
                reloadSach();
                TinhTongSach();
            }
        }

        private void btnSuaSach_Click(object sender, EventArgs e)
        {
            btnThemSach.Enabled = false;
            btnXoaSach.Enabled = false;
            btnUpSach.Enabled = true;

            btnThemSach.BackColor = Color.PeachPuff;
            btnXoaSach.BackColor = Color.PeachPuff;
            btnUpSach.BackColor = Color.Peru;
        }

        private void btnUpSach_Click(object sender, EventArgs e)
        {
            if (txtMaSach.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã sách!\nVui lòng nhập mã sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaSach.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from SACH where MASACH='" + txtMaSach.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaSach.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSach.Focus();
                    ResetValues_Sach();
                    return;
                }
                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (txtTenSach.Text.Trim() != "" && txtTenSach.Text != tbDanhSach.Rows[0]["TENSACH"].ToString())
                    {
                        string sqlUpdateTenSach = "update SACH set TENSACH=(N'" + txtTenSach.Text + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateTenSach);
                    }
                    if (cmbTacGia.SelectedIndex != -1 && cmbTacGia.SelectedValue.ToString() != tbDanhSach.Rows[0]["MATACGIA"].ToString())
                    {
                        string sqlUpdateMaTG = "update SACH set MATACGIA=(N'" + cmbTacGia.SelectedValue.ToString() + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateMaTG);
                    }
                    if (cmbTheLoai.SelectedIndex != -1 && cmbTheLoai.SelectedValue.ToString() != tbDanhSach.Rows[0]["MATHELOAI"].ToString())
                    {
                        string sqlUpdateMaTL = "update SACH set MATHELOAI=(N'" + cmbTheLoai.SelectedValue.ToString() + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateMaTL);
                    }
                    if (cmbNXB.SelectedIndex != -1 && cmbNXB.SelectedValue.ToString() != tbDanhSach.Rows[0]["MA_NXB"].ToString())
                    {
                        string sqlUpdateMaNXB = "update SACH set MA_NXB=(N'" + cmbNXB.SelectedValue.ToString() + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateMaNXB);
                    }
                    if (txtSLSach.Text.Trim() != "" && txtSLSach.Text != tbDanhSach.Rows[0]["TONGSOLUONG"].ToString())
                    {
                        string sqlUpdateSLSach = "update SACH set TONGSOLUONG=(N'" + txtSLSach.Text + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateSLSach);
                    }
                    if (txtGiaBia.Text.Trim() != "" && txtGiaBia.Text != tbDanhSach.Rows[0]["GIABIA"].ToString())
                    {
                        string sqlUpdateSLSach = "update SACH set GIABIA=(N'" + txtGiaBia.Text + "') where MASACH = (N'" + txtMaSach.Text + "')";
                        dtBase.Updatedate(sqlUpdateSLSach);
                    }
                    LoadDataSach();
                    ResetValues_Sach();
                    reloadSach();
                }
            }
        }
        private void btnXoaSach_Click(object sender, EventArgs e)
        {
            if (txtMaSach.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã sách!\nVui lòng nhập mã sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaSach.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from SACH where MASACH='" + txtMaSach.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaSach.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSach.Focus();
                    ResetValues_Sach();
                    return;
                }
                DataTable tbSPM = dtBase.DataSelect("select * from PHIEUMUON where PHIEUMUON.MASACH = '" + txtMaSach.Text + "'");
                if(tbSPM.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn không thể xóa sách này.\nLý do: Sinh viên đang mượn loại sách này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaSach.Focus();
                    ResetValues_Sach();
                    return;
                }
                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDeleteSach = "delete from SACH where MASACH = (N'" + txtMaSach.Text + "')";
                    dtBase.Updatedate(sqlDeleteSach);
                    LoadDataSach();
                    ResetValues_Sach();
                    loadCmbSach();
                    cmbSach.SelectedIndex = -1;
                    TinhTongSach();
                }
            }
        }
        private void txtTKSach_TextChanged(object sender, EventArgs e)
        {
            dgvSach.DataSource = dtBase.DataSelect("Select SACH.MASACH, SACH.TENSACH, TACGIA.TENTG, NXB.TENNXB, THELOAI.TENTHELOAI, SACH.TONGSOLUONG, GIABIA " +
                "from ((SACH left join TACGIA ON SACH.MATACGIA = TACGIA.MATACGIA) left join THELOAI ON SACH.MATHELOAI = THELOAI.MATHELOAI) left join NXB ON SACH.MA_NXB = NXB.MA_NXB " +
                "where MASACH like (N'%" + txtTKSach.Text + "%') or TENSACH like (N'%" + txtTKSach.Text + "%')");
        }

        ///Tac gia
        ///
        public void ResetValues_TG()
        {
            txtMaTacGia.Text = "";
            txtTenTacGia.Text = "";
        }
        void reloadTG()
        {
            txtMaTacGia.Text = "";
            txtMaTacGia.Enabled = true;
            btnThemTG.Enabled = true;
            btnSuaTG.Enabled = true;
            btnXoaTG.Enabled = true;
            btnLuuTG.Enabled = false;
            btnUpTG.Enabled = false;

            btnThemTG.BackColor = Color.Peru;
            btnSuaTG.BackColor = Color.Peru;
            btnXoaTG.BackColor = Color.Peru;
            btnLuuTG.BackColor = Color.PeachPuff;
            btnUpTG.BackColor = Color.PeachPuff;

        }
        private void btnReloadTG_Click(object sender, EventArgs e)
        {
            txtTKTG.Text = "";
            ResetValues_TG();
            LoadDataTG();
            reloadTG();
        }
        private void btnThemTG_Click(object sender, EventArgs e)
        {
            txtMaTacGia.Text = sinhma.SinhMa("TACGIA", "TG", "MATACGIA");
            txtMaTacGia.Enabled = false;
            txtTenTacGia.Text = "";

            btnSuaTG.Enabled = false;
            btnXoaTG.Enabled = false;
            btnLuuTG.Enabled = true;

            btnSuaTG.BackColor = Color.PeachPuff;
            btnXoaTG.BackColor = Color.PeachPuff;
            btnLuuTG.BackColor = Color.Peru;
        }
        private void btnLuuTG_Click(object sender, EventArgs e)
        {
            if (txtMaTacGia.Text.Trim() == "" || txtTenTacGia.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin trước khi thêm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTacGia.Focus();
                return;
            }
            else
            {
                string sqlInsert = "insert into TACGIA values(N'" + txtMaTacGia.Text + "', N'" + txtTenTacGia.Text + "')";
                dtBase.Updatedate(sqlInsert);
                MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDataTG();
                ResetValues_TG();
                loadCmbTacGia();
                cmbTacGia.SelectedIndex = -1;
                reloadTG();
            }
        }

        private void btnSuaTG_Click(object sender, EventArgs e)
        {
            btnThemTG.Enabled = false;
            btnXoaTG.Enabled = false;
            btnUpTG.Enabled = true;

            btnThemTG.BackColor = Color.PeachPuff;
            btnXoaTG.BackColor = Color.PeachPuff;
            btnUpTG.BackColor = Color.Peru;

        }
        private void btnUpTG_Click(object sender, EventArgs e)
        {
            if (txtMaTacGia.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã tác giả!\nVui lòng nhập mã tác giả! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTacGia.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from TACGIA where MATACGIA='" + txtMaTacGia.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaTacGia.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTacGia.Focus();
                    ResetValues_TG();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (txtTenTacGia.Text != "" && txtTenTacGia.Text != tbDanhSach.Rows[0]["TENTG"].ToString())
                    {
                        string sqlUpdateTenTG = "update TACGIA set TENTG = (N'" + txtTenTacGia.Text + "') where MATACGIA =('" + txtMaTacGia.Text + "')";
                        dtBase.Updatedate(sqlUpdateTenTG);
                    }
                    LoadDataTG();
                    ResetValues_TG();
                    loadCmbTacGia();
                    cmbTacGia.SelectedIndex = -1;
                    reloadTG();
                }
            }
        }
        private void btnXoaTG_Click(object sender, EventArgs e)
        {
            if (txtMaTacGia.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã tác giả!\nVui lòng nhập mã tác giả! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTacGia.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from TACGIA where MATACGIA='" + txtMaTacGia.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaTacGia.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTacGia.Focus();
                    ResetValues_TG();
                    return;
                }

                DataTable tbTG_S = dtBase.DataSelect("select * from SACH where SACH.MATACGIA = '" + txtMaTacGia.Text + "'");
                if (tbTG_S.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn không thể xóa tác giả này!\nLý do: Trong kho đang có sách do tác giả này viết!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTacGia.Focus();
                    ResetValues_TG();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDeleteTG = "delete from TACGIA where MATACGIA = (N'" + txtMaTacGia.Text + "')";
                    dtBase.Updatedate(sqlDeleteTG);
                    LoadDataTG();
                    ResetValues_TG();
                    loadCmbTacGia();
                    cmbTacGia.SelectedIndex = -1;
                }
            }
        }
        private void txtTKTG_TextChanged(object sender, EventArgs e)
        {
            dgvTacGia.DataSource = dtBase.DataSelect("Select * from TACGIA where MATACGIA like (N'%" + txtTKTG.Text + "%') or TENTG like (N'%" + txtTKTG.Text + "%')");
        }

        ///thể loại
        public void ResetValueTL()
        {
            txtMaTheLoai.Text = "";
            txtTenTheLoai.Text = "";
        }
        void reloadTL()
        {
            txtMaTheLoai.Text = "";
            txtMaTheLoai.Enabled = true;

            btnThemTheLoai.Enabled = true;
            btnSuaTheLoai.Enabled = true;
            btnXoaTheLoai.Enabled = true;
            btnLuuTL.Enabled = false;
            btnUpTL.Enabled = false;

            btnThemTheLoai.BackColor = Color.Peru;
            btnSuaTheLoai.BackColor = Color.Peru;
            btnXoaTheLoai.BackColor = Color.Peru;
            btnLuuTL.BackColor = Color.PeachPuff;
            btnUpTL.BackColor = Color.PeachPuff;

        }
        private void btnReloadTL_Click(object sender, EventArgs e)
        {
            txtTKTL.Text = "";
            ResetValueTL();
            LoadDataTheLoai();
            reloadTL();
        }
        private void btnThemTheLoai_Click(object sender, EventArgs e)
        {
            txtMaTheLoai.Text = sinhma.SinhMa("THELOAI", "TL", "MATHELOAI");
            txtMaTheLoai.Enabled = false;

            txtTenTheLoai.Text = "";

            btnSuaTheLoai.Enabled = false;
            btnXoaTheLoai.Enabled = false;
            btnLuuTL.Enabled = true;

            btnSuaTheLoai.BackColor = Color.PeachPuff;
            btnXoaTheLoai.BackColor = Color.PeachPuff;
            btnLuuTL.BackColor = Color.Peru;

        }
        private void btnLuuTL_Click(object sender, EventArgs e)
        {
            if (txtMaTheLoai.Text.Trim() == "" || txtTenTheLoai.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin trước khi thêm thể loại! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTheLoai.Focus();
                return;
            }
            else
            {
                string sqlInsert = "insert into THELOAI values(N'" + txtMaTheLoai.Text + "', N'" + txtTenTheLoai.Text + "')";
                dtBase.Updatedate(sqlInsert);
                MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDataTheLoai();
                loadCmbTheLoai();
                cmbTheLoai.SelectedIndex = -1;
                ResetValueTL();
                reloadTL();
            }
        }

        private void btnSuaTheLoai_Click(object sender, EventArgs e)
        {
            btnThemTheLoai.Enabled = false;
            btnXoaTheLoai.Enabled = false;
            btnUpTL.Enabled = true;

            btnThemTheLoai.BackColor = Color.PeachPuff;
            btnXoaTheLoai.BackColor = Color.PeachPuff;
            btnUpTL.BackColor = Color.Peru;
        }

        private void btnUpTL_Click(object sender, EventArgs e)
        {
            if (txtMaTheLoai.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã thể loại!\nVui lòng nhập mã thể loại! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTheLoai.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from THELOAI where MATHELOAI='" + txtMaTheLoai.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaTheLoai.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTheLoai.Focus();
                    ResetValueTL();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (txtTenTheLoai.Text != "" && txtTenTheLoai.Text != tbDanhSach.Rows[0]["TENTHELOAI"].ToString())
                    {
                        string sqlUpdateTenTL = "update THELOAI set TENTHELOAI = (N'" + txtTenTheLoai.Text + "') where MATHELOAI = '" + txtMaTheLoai.Text + "'";
                        dtBase.Updatedate(sqlUpdateTenTL);
                    }
                    LoadDataTheLoai();
                    ResetValueTL();
                    reloadTL();
                    loadCmbTheLoai();
                    cmbTheLoai.SelectedIndex = -1;
                }
            }
        }
        private void btnXoaTheLoai_Click(object sender, EventArgs e)
        {
            if (txtMaTheLoai.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã thể loại!\nVui lòng nhập mã thể loại! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaTheLoai.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from THELOAI where MATHELOAI='" + txtMaTheLoai.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaTheLoai.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTheLoai.Focus();
                    ResetValueTL();
                    return;
                }

                DataTable tbTL_S = dtBase.DataSelect("select * from SACH where SACH.MATHELOAI = '" + txtMaTheLoai.Text + "'");
                if (tbTL_S.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn không thể xóa thể loại này!\nLý do: Trong kho đang có sách thuộc thể loại này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaTheLoai.Focus();
                    ResetValueTL();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDeleteTL = "delete from THELOAI where MATHELOAI = (N'" + txtMaTheLoai.Text + "')";
                    dtBase.Updatedate(sqlDeleteTL);
                    LoadDataTheLoai();
                    ResetValueTL();
                    loadCmbTheLoai();
                    cmbTheLoai.SelectedIndex = -1;
                }
            }
        }
        private void txtTKTL_TextChanged(object sender, EventArgs e)
        {
            dgvTheLoai.DataSource = dtBase.DataSelect("Select * from THELOAI where MATHELOAI like (N'%" + txtTKTL.Text + "%') or TENTHELOAI like (N'%" + txtTKTL.Text + "%')");
        }

        ///NXB
        ///
        public void ResetValueNXB()
        {
            txtMaNXBan.Text = "";
            txtTenNXB.Text = "";
            txtDiaChiNXB.Text = "";
            txtSDT_NXB.Text = "";
        }
        void reloadNXB()
        {
            txtMaNXBan.Text = "";
            txtMaNXBan.Enabled = true;

            btnThemNXB.Enabled = true;
            btbSuaNXB.Enabled = true;
            btnXoaNXB.Enabled = true;
            btnLuuNXB.Enabled = false;
            btnUpNXB.Enabled = false;

            btnThemNXB.BackColor = Color.Peru;
            btbSuaNXB.BackColor = Color.Peru;
            btnXoaNXB.BackColor = Color.Peru;
            btnLuuNXB.BackColor = Color.PeachPuff;
            btnUpNXB.BackColor = Color.PeachPuff;

        }
        private void btnReloadNXB_Click(object sender, EventArgs e)
        {
            txtTKNXB.Text = "";
            ResetValueNXB();
            LoadDataNXB();
            reloadNXB();
        }
        private void btnThemNXB_Click(object sender, EventArgs e)
        {
            txtMaNXBan.Text = sinhma.SinhMa("NXB", "XB", "MA_NXB");
            txtMaNXBan.Enabled = false;

            txtTenNXB.Text = "";
            txtDiaChiNXB.Text = "";
            txtSDT_NXB.Text = "";

            btbSuaNXB.Enabled = false;
            btnXoaNXB.Enabled = false;
            btnLuuNXB.Enabled = true;

            btbSuaNXB.BackColor = Color.PeachPuff;
            btnXoaNXB.BackColor = Color.PeachPuff;
            btnLuuNXB.BackColor = Color.Peru;

        }
        private void btnLuuNXB_Click(object sender, EventArgs e)
        {
            if (txtMaNXBan.Text.Trim() == "" || txtTenNXB.Text.Trim() == "" || txtDiaChiNXB.Text.Trim() == "" || txtSDT_NXB.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin trước khi thêm NXB! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaNXBan.Focus();
                return;
            }
            if (txtSDT_NXB.Text.Length != 10 && txtSDT_NXB.Text.Length != 11)
            {
                MessageBox.Show("SĐT không đúng định dạng!\nVui lòng nhập lại!\n(SĐT phải có đủ 10 hoặc 11 chữ số)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSDT.Focus();
                return;
            }
            if (txtSDT_NXB.Text.StartsWith("84") == false && txtSDT_NXB.Text.StartsWith("0") == false)
            {
                MessageBox.Show("SĐT không đúng định dạng!\nVui lòng nhập lại!\n(SĐT phải đầu số là (84) hoặc (0)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSDT.Focus();
                return;
            }
            else
            {
                string sqlInsert = "insert into NXB values(N'" + txtMaNXBan.Text + "', N'" + txtTenNXB.Text + "', N'" + txtDiaChiNXB.Text + "', N'" + txtSDT_NXB.Text + "')";
                dtBase.Updatedate(sqlInsert);
                MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDataNXB();
                ResetValueNXB();
                loadCmbNXB();
                cmbNXB.SelectedIndex = -1;
                reloadNXB();
            }
        }

        private void btbSuaNXB_Click(object sender, EventArgs e)
        {
            btnThemNXB.Enabled = false;
            btnXoaNXB.Enabled = false;
            btnUpNXB.Enabled = true;

            btnThemNXB.BackColor = Color.PeachPuff;
            btnXoaNXB.BackColor = Color.PeachPuff;
            btnUpNXB.BackColor = Color.Peru;

        }

        private void btnUpNXB_Click(object sender, EventArgs e)
        {
            if (txtMaNXBan.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã NXB!\nVui lòng nhập mã NXB! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaNXBan.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from NXB where MA_NXB='" + txtMaNXBan.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaNXBan.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaNXBan.Focus();
                    ResetValueNXB();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (txtTenNXB.Text != "" && txtTenNXB.Text != tbDanhSach.Rows[0]["TENNXB"].ToString())
                    {
                        string sqlUpdateTenNXB = "update NXB set TENNXB = (N'" + txtTenNXB.Text + "') where MA_NXB='" + txtMaNXBan.Text + "'";
                        dtBase.Updatedate(sqlUpdateTenNXB);
                    }
                    if (txtDiaChiNXB.Text != "" && txtDiaChiNXB.Text != tbDanhSach.Rows[0]["DIACHI_NXB"].ToString())
                    {
                        string sqlUpdateTenNXB = "update NXB set DIACHI_NXB = (N'" + txtDiaChiNXB.Text + "') where MA_NXB='" + txtMaNXBan.Text + "'";
                        dtBase.Updatedate(sqlUpdateTenNXB);
                    }
                    if (txtSDT_NXB.Text != "" && txtSDT_NXB.Text != tbDanhSach.Rows[0]["SDT_NXB"].ToString())
                    {
                        string sqlUpdateTenNXB = "update NXB set SDT_NXB = (N'" + txtSDT_NXB.Text + "') where MA_NXB='" + txtMaNXBan.Text + "'";
                        dtBase.Updatedate(sqlUpdateTenNXB);
                    }
                    LoadDataNXB();
                    ResetValueNXB();
                    reloadNXB();
                    loadCmbNXB();
                    cmbNXB.SelectedIndex = -1;
                }
            }
        }
        private void btnXoaNXB_Click(object sender, EventArgs e)
        {
            if (txtMaNXBan.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã NXB!\nVui lòng nhập mã NXB! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaNXBan.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from NXB where MA_NXB='" + txtMaNXBan.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaNXBan.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaNXBan.Focus();
                    ResetValueNXB();
                    return;
                }

                DataTable tbNXB_S = dtBase.DataSelect("select * from SACH where SACH.MA_NXB = '" + txtMaNXBan.Text + "'");
                if (tbNXB_S.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn không thể xóa NXB này!\nLý do: Trong kho đang có sách do NXB này xuất bản!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaNXBan.Focus();
                    ResetValueNXB();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDeleteNXB = "delete from NXB where MA_NXB = '" + txtMaNXBan.Text + "'";
                    dtBase.Updatedate(sqlDeleteNXB);
                    LoadDataNXB();
                    ResetValueNXB();
                    loadCmbNXB();
                    cmbNXB.SelectedIndex = -1;
                }
            }
        }
        private void txtTKNXB_TextChanged(object sender, EventArgs e)
        {
            dgvNXB.DataSource = dtBase.DataSelect("Select * from NXB where MA_NXB like (N'%" + txtTKNXB.Text + "%') or TENNXB like (N'%" + txtTKNXB.Text + "%')");
        }

        ///PHIEU MUON
        ///
        private void cmbThang_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvPhieuMuon.DataSource = dtBase.DataSelect("Select PHIEUMUON.MAPHIEUMUON, PHIEUMUON.MASV, HOTEN, TENSACH, PHIEUMUON.SOLUONGMUON, PHIEUMUON.NGAYMUON, PHIEUMUON.NGAYHENTRA, PHIEUMUON.TRANGTHAI " +
                "from (PHIEUMUON left join SinhVien on PHIEUMUON.MASV = SinhVien.MASV) left join SACH on PHIEUMUON.MASACH = SACH.MASACH " +
                "where month(PHIEUMUON.NGAYMUON)= '" + cmbThang.Text + "' and  year(PHIEUMUON.NGAYMUON)= '" + cmbNam.Text + "'");
        }
        private void cmbLocPM_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cmbLocPhieuMuon = "Select PHIEUMUON.MAPHIEUMUON, PHIEUMUON.MASV, HOTEN, TENSACH, " +
               "PHIEUMUON.SOLUONGMUON, PHIEUMUON.NGAYMUON, PHIEUMUON.NGAYHENTRA, PHIEUMUON.TRANGTHAI " +
               "from (PHIEUMUON left join SinhVien on PHIEUMUON.MASV = SinhVien.MASV) left join SACH on PHIEUMUON.MASACH = SACH.MASACH ";

            if (cmbLocPM.Text == "Tất cả")
            {
                dgvPhieuMuon.DataSource = dtBase.DataSelect(cmbLocPhieuMuon);
            }
            else if (cmbLocPM.Text == "Đã trả")
            {
                dgvPhieuMuon.DataSource = dtBase.DataSelect(cmbLocPhieuMuon + " where PHIEUMUON.TRANGTHAI = N'Đã trả'");
            }
            else if (cmbLocPM.Text == "Đang mượn")
            {
                dgvPhieuMuon.DataSource = dtBase.DataSelect(cmbLocPhieuMuon + " where PHIEUMUON.TRANGTHAI= N'Đang mượn sách'");
            }
            else if (cmbLocPM.Text == "Quá hạn")
            {
                dgvPhieuMuon.DataSource = dtBase.DataSelect(cmbLocPhieuMuon + " where PHIEUMUON.TRANGTHAI= N'Quá hạn'");
            }
            else
            {
                dgvPhieuMuon.DataSource = dtBase.DataSelect(cmbLocPhieuMuon + " where PHIEUMUON.TRANGTHAI like N'%Không trả sách%'");
            }
        }
        public void ResetvaluePM()
        {
            txtMaPM.Text = "";
            txtMSV.Text = "";
            txtHoTenPM.Text = "";
            cmbSach.SelectedIndex = -1;
            txtSLMuon.Text = "";
            cmbTrangThai.SelectedIndex = -1;
            cmbTrangThai.Text = "";
            dtpNgayMuon.Text = now.ToString();
            dtpNgayHenTra.Text = now.ToString();
            cmbThang.Text = now.Month.ToString();
            cmbNam.Text = now.Year.ToString();
            cmbLocPM.Text = "";
        }
        private void cmbSach_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSach.Text != "")
            {
                txtSLMuon.Text = "1";
            }
            else
            {
                txtSLMuon.Text = "";
            }
        }

        private void txtMSV_TextChanged(object sender, EventArgs e)
        {
            if (txtMSV.Text.Trim() != "")
            {
                DataTable tbSVPM = dtBase.DataSelect("select * from SINHVIEN where SINHVIEN.MASV = '" + txtMSV.Text + "'");
                if (tbSVPM.Rows.Count > 0)
                {
                    txtHoTenPM.Text = tbSVPM.Rows[0]["HOTEN"].ToString();
                }
            }
            else
            {
                txtHoTenPM.Text = "";
            }
        }
        private void reloadPM()
        {
            txtMaPM.Enabled = true;
            txtMSV.Enabled = true;
            dtpNgayMuon.Enabled = true;
            cmbSach.Enabled = true;
            btnTaoPhieu.Enabled = true;
            btnXoaPhieuMuon.Enabled = true;
            btnSuaPhieuMuon.Enabled = true;
            btnLuuPM.Enabled = false;
            btnUpPM.Enabled = false;

            btnTaoPhieu.BackColor = Color.Peru;
            btnXoaPhieuMuon.BackColor = Color.Peru;
            btnSuaPhieuMuon.BackColor = Color.Peru;
            btnLuuPM.BackColor = Color.PeachPuff;
            btnUpPM.BackColor = Color.PeachPuff;

        }
        private void btnThemMoi_Click(object sender, EventArgs e)
        {
            txtTKPM.Text = "";
            ResetvaluePM();
            LoadDataPhieuMuon();
            reloadPM();
        }
        private void btnTaoPhieu_Click(object sender, EventArgs e)
        {
            txtMaPM.Text = sinhma.SinhMa("PHIEUMUON", "PM", "MAPHIEUMUON");
            txtMaPM.Enabled = false;

            txtMSV.Text = "";
            txtHoTenPM.Text = "";
            cmbSach.SelectedIndex = -1;
            txtSLMuon.Text = "";
            cmbTrangThai.SelectedIndex = -1;
            cmbTrangThai.Text = "";
            dtpNgayMuon.Text = now.ToString();
            dtpNgayHenTra.Text = now.ToString();

            btnLuuPM.Enabled = true;
            btnXoaPhieuMuon.Enabled = false;
            btnSuaPhieuMuon.Enabled = false;

            btnLuuPM.BackColor = Color.Peru;
            btnXoaPhieuMuon.BackColor = Color.PeachPuff;
            btnSuaPhieuMuon.BackColor = Color.PeachPuff;

        }

        private void btnLuuPM_Click(object sender, EventArgs e)
        {
            if (txtMaPM.Text.Trim() == "" || txtMSV.Text.Trim() == "" || cmbSach.Text.Trim() == "" || cmbTrangThai.SelectedIndex == -1 || cmbTrangThai.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin trước khi thêm phiếu mượn! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPM.Focus();
                return;
            }
            else
            {
                DataTable tbTaiKhoanBiKhoa = dtBase.DataSelect("select * from TAIKHOAN_BIKHOA where MASV_BIKHOA ='" + txtMSV.Text + "'");
                if (tbTaiKhoanBiKhoa.Rows.Count == 1)
                {
                    MessageBox.Show("Tai khoản này đã bị khóa!\nBạn không thể tạo phiếu mượn cho tài khoản này\nLý do: Không trả sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataTable PM_SACH = dtBase.DataSelect("select * from PHIEUMUON where MASV = '" + txtMSV.Text + "' and (TRANGTHAI = N'Đang mượn sách' or TRANGTHAI = N'Quá hạn')");
                if(PM_SACH.Rows.Count == 3)
                {
                    MessageBox.Show("Sinh viên chỉ được mượn tối đa 3 quyển sách!\nBạn không thế tạo phiếu mượn này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ResetvaluePM();
                    return;
                }

                DataTable PM_MaSACH = dtBase.DataSelect("select * from PHIEUMUON join SACH on PHIEUMUON.MASACH = SACH.MASACH where PHIEUMUON.MASV = '" + txtMSV.Text + "' and PHIEUMUON.MASACH = '"+ cmbSach.SelectedValue.ToString() + "' and (PHIEUMUON.TRANGTHAI = N'Đang mượn sách' or PHIEUMUON.TRANGTHAI = N'Quá hạn')");
                if (PM_MaSACH.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn đang mượn sách '" + PM_MaSACH.Rows[0]["TENSACH"].ToString() + "' nên bạn không thể mượn thêm loại sách này!\nMỗi sinh viên chỉ được mượn một quyển sách cùng loại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if(dtpNgayMuon.Value.ToString().Substring(0, 11) != now.ToString().Substring(0, 11))
                {
                    MessageBox.Show("Ngày tạo phiếu mượn phải là ngày hôm nay '"+ dtpNgayMuon.Value.ToString().Substring(0, 11) + "', '"+now.ToString().Substring(0, 11) + "'!\n Bạn không được thay đổi ngày mượn sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if(dtpNgayHenTra.Value <= dtpNgayMuon.Value)
                {
                    MessageBox.Show("Ngày hẹn trả phải lớn hơn ngày mượn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                DataTable tbSach = dtBase.DataSelect("select * from SACH where MASACH = '" + cmbSach.SelectedValue.ToString() + "'");
                DataTable tbTaiKhoanSV = dtBase.DataSelect("select * from SINHVIEN where MASV='" + txtMSV.Text + "'");
                int TongSach;
                if (tbSach.Rows.Count > 0)
                {
                    TongSach = int.Parse(tbSach.Rows[0]["TongSoLuong"].ToString());
                    if (tbTaiKhoanSV.Rows.Count == 1 && TongSach > 0)
                    {
                        string sqlInsert = "insert into PHIEUMUON values(N'" + txtMaPM.Text + "', N'" + txtMSV.Text + "', N'" + cmbSach.SelectedValue.ToString() + "', N'" + txtSLMuon.Text + "', N'" + dtpNgayMuon.Value.ToString("yyyy-MM-dd") + "', N'" + dtpNgayHenTra.Value.ToString("yyyy-MM-dd") + "', N'" + cmbTrangThai.Text + "')";
                        dtBase.Updatedate(sqlInsert);

                        TongSach = TongSach - 1;
                        string sqlUpdateSLSach = "update SACH set SACH.TONGSOLUONG = '" + TongSach.ToString() + "' where SACH.MASACH = '" + cmbSach.SelectedValue.ToString() + "'";
                        dtBase.Updatedate(sqlUpdateSLSach);
                        LoadDataSach();

                        MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDataPhieuMuon();
                        ResetvaluePM();
                        TinhTongPM();
                        reloadPM();

                    }
                    else if (tbTaiKhoanSV.Rows.Count == 0)
                    {
                        MessageBox.Show("Không tồn tại mã sinh viên này!\nVui lòng kiểm tra lại mã sinh viên của bạn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtMSV.Focus();
                        return;
                    }
                    else if (TongSach == 0)
                    {
                        if (MessageBox.Show("Loại sách này đã hết.\nBạn không thể mượn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                        {
                            return;
                        }
                    }
                }
            }
        }

        private void btnSuaPhieuMuon_Click(object sender, EventArgs e)
        {
            txtMSV.Enabled = false;
            dtpNgayMuon.Enabled = false;
            cmbSach.Enabled = false;
            btnTaoPhieu.Enabled = false;
            btnXoaPhieuMuon.Enabled = false;
            btnUpPM.Enabled = true;

            btnTaoPhieu.BackColor = Color.PeachPuff;
            btnXoaPhieuMuon.BackColor = Color.PeachPuff;
            btnUpPM.BackColor = Color.Peru;
        }

        private void btnUpPM_Click(object sender, EventArgs e)
        {
            if (txtMaPM.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã phiếu mượn!\nVui lòng nhập mã phiếu mượn để tiến hành sửa! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPM.Focus();
                return;
            }
            else
            {
                /// lấy ra dữ liệu của phiếu mượn định sửa
                DataTable tbDanhSach = dtBase.DataSelect("select * from PHIEUMUON where MAPHIEUMUON = '" + txtMaPM.Text + "'");

                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaPM.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaPM.Focus();
                    ResetvaluePM();
                    return;
                }
                
                if (tbDanhSach.Rows.Count > 0)
                {
                    /*
                    string masach;
                    masach = tbDanhSach.Rows[0]["MASACH"].ToString();

                    ///lấy ra dữ liệu của sách tưởng ứng với phiếu mượn định sửa
                    DataTable tbSach_old = dtBase.DataSelect("select * from SACH where MASACH = '" + masach + "'");
                    int TongSach_old;
                    TongSach_old = int.Parse(tbSach_old.Rows[0]["TongSoLuong"].ToString());

                    ///Tongsach: số lượng sách hiện tại của loại sách vừa được thay thế
                    DataTable tbSach = dtBase.DataSelect("select * from SACH where MASACH = '" + cmbSach.SelectedValue.ToString() + "'");
                    int TongSach;
                    TongSach = int.Parse(tbSach.Rows[0]["TongSoLuong"].ToString());
                    */
                    if (tbDanhSach.Rows[0]["TRANGTHAI"].ToString() == "Đã trả")
                    {
                        MessageBox.Show("Người mượn sách đã trả sách cho phiếu mượn này!\nBạn không được sửa thông tin của phiếu này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ResetvaluePM();
                        cmbTrangThai.Text = "";
                        return;
                    }
                    if (tbDanhSach.Rows[0]["TRANGTHAI"].ToString() == "Không trả sách, nhân viên đã lập phiếu trả")
                    {
                        MessageBox.Show("Nhân viên đã lập phiếu trả cho phiếu mượn này!\nBạn không được sửa thông tin của phiếu này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ResetvaluePM();
                        cmbTrangThai.Text = "";
                        return;
                    }
                    /*
                    DataTable PM_MaSACH = dtBase.DataSelect("select * from PHIEUMUON where MASV = '" + txtMSV.Text + "' and MASACH = '" + cmbSach.SelectedValue.ToString() + "' and TRANGTHAI = N'Đang mượn sách'");
                    if (PM_MaSACH.Rows.Count > 0)
                    {
                        MessageBox.Show("Bạn đang mượn sách '" + PM_MaSACH.Rows[0]["TENSACH"].ToString() + "' nên bạn không thể mượn thêm loại sách này!\nMỗi sinh viên chỉ được mượn một quyển sách cùng loại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (TongSach == 0)
                    {
                        MessageBox.Show("Loại sách này đã hết!\nBạn không thể mượn sách này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }*/
                    if (dtpNgayHenTra.Value <= dtpNgayMuon.Value || dtpNgayHenTra.Value <= now)
                    {
                        MessageBox.Show("Ngày hẹn trả phải lớn hơn ngày mượn và ngày hôm nay!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if(cmbTrangThai.Text != "Đang mượn sách")
                    {
                        MessageBox.Show("Bạn chưa cập nhật trạng thái mượn!\nVui lòng cập nhật trạng thái mượn để hoàn thành gia hạn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        cmbTrangThai.Focus();
                        return;
                    }
                    if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        /*
                        if (cmbSach.SelectedValue.ToString() != tbDanhSach.Rows[0]["MASACH"].ToString() && cmbSach.SelectedIndex != -1)
                        {
                            string sqlUpdateMS = "update PHIEUMUON set MASACH = (N'" + cmbSach.SelectedValue.ToString() + "') where MAPHIEUMUON='" + txtMaPM.Text + "'";
                            dtBase.Updatedate(sqlUpdateMS);

                            ///update số lượng sách mới vừa được thay thế 
                            TongSach = TongSach - 1;
                            string sqlUpdateSLSach = "update SACH set SACH.TONGSOLUONG = '" + TongSach.ToString() + "' where SACH.MASACH = '" + cmbSach.SelectedValue.ToString() + "'";
                            dtBase.Updatedate(sqlUpdateSLSach);
                            LoadDataSach();

                            ///update số lượng sách cũ vừa bị thay thế
                            TongSach_old = TongSach_old + 1;
                            string sqlUpdateSLSach_old = "update SACH set SACH.TONGSOLUONG = '" + TongSach_old.ToString() + "' where SACH.MASACH = '" + masach + "'";
                            dtBase.Updatedate(sqlUpdateSLSach_old);
                            LoadDataSach();

                        }
                        */
                        if (dtpNgayHenTra.Text != tbDanhSach.Rows[0]["NGAYHENTRA"].ToString() && dtpNgayHenTra.Text.Trim() != "")
                        {
                            string sqlUpdateHen = "update PHIEUMUON set NGAYHENTRA = (N'" + dtpNgayHenTra.Value.ToString("yyyy-MM-dd") + "') where MAPHIEUMUON='" + txtMaPM.Text + "'";
                            dtBase.Updatedate(sqlUpdateHen);
                        }
                        if (cmbTrangThai.Text != tbDanhSach.Rows[0]["TRANGTHAI"].ToString() && cmbTrangThai.Text.Trim() != "")
                        {
                            string sqlUpdateTT= "update PHIEUMUON set TRANGTHAI = (N'" + cmbTrangThai.Text + "') where MAPHIEUMUON='" + txtMaPM.Text + "'";
                            dtBase.Updatedate(sqlUpdateTT);
                        }
                        LoadDataPhieuMuon();
                        ResetvaluePM();
                        reloadPM();
                    }
                }
            }
        }
        private void btnXoaPhieuMuon_Click(object sender, EventArgs e)
        {
            if (txtMaPM.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã phiếu mượn!\nVui lòng nhập mã phiếu mượn để tiến hành xóa! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPM.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from PHIEUMUON where MAPHIEUMUON='" + txtMaPM.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaPM.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaPM.Focus();
                    ResetvaluePM();
                    return;
                }
                DataTable tbDS_Xoa = dtBase.DataSelect("select * from PHIEUMUON where MAPHIEUMUON = '" + txtMaPM.Text + "' and (TRANGTHAI = N'Đang mượn sách' or TRANGTHAI = N'Không trả sách' or TRANGTHAI = N'Quá hạn')");
                if(tbDS_Xoa.Rows.Count > 0)
                {
                    MessageBox.Show("Bạn không thể xóa phiếu mượn này!\nLý do: Sinh viên đang mượn sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaPM.Focus();
                    ResetvaluePM();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sqlDeletePM = "delete from PHIEUMUON where MAPHIEUMUON = (N'" + txtMaPM.Text + "')";
                    dtBase.Updatedate(sqlDeletePM);
                    LoadDataPhieuMuon();
                    TinhTongPM();
                    ResetvaluePM();
                    cmbTrangThai.SelectedIndex = -1;
                }
            }
        }

        private void txtTKPM_TextChanged(object sender, EventArgs e)
        {
            dgvPhieuMuon.DataSource = dtBase.DataSelect("Select PHIEUMUON.MAPHIEUMUON, PHIEUMUON.MASV, SinhVien.HOTEN, SACH.TENSACH, " +
               "PHIEUMUON.SOLUONGMUON, PHIEUMUON.NGAYMUON, PHIEUMUON.NGAYHENTRA, PHIEUMUON.TRANGTHAI " +
               "from (PHIEUMUON left join SinhVien on PHIEUMUON.MASV = SinhVien.MASV) left join SACH on PHIEUMUON.MASACH = SACH.MASACH" +
               " where PHIEUMUON.MAPHIEUMUON like(N'%" + txtTKPM.Text + "%') or PHIEUMUON.MASV like(N'%" + txtTKPM.Text + "%')");
        }

        ///phiếu trả
        //
        private void txtMSVPT_TextChanged(object sender, EventArgs e)
        {
            if (txtMSVPT.Text.Trim() != "")
            {
                cmbSachPT.DataSource = dtBase.DataSelect("select SACH.TENSACH, SACH.MASACH from SACH join PHIEUMUON on PHIEUMUON.MASACH = SACH.MASACH where PHIEUMUON.MASV = '" + txtMSVPT.Text + "' " +
                    "and (PHIEUMUON.TRANGTHAI = N'Đang mượn sách' or PHIEUMUON.TRANGTHAI = N'Quá hạn' or PHIEUMUON.TRANGTHAI = N'Không trả sách' )");
                cmbSachPT.DisplayMember = "TENSACH";
                cmbSachPT.ValueMember = "MASACH";

                ////
                DataTable tbSVPT = dtBase.DataSelect("select * from SINHVIEN where MASV = '" + txtMSVPT.Text + "'");
                if(tbSVPT.Rows.Count > 0)
                {
                    txtHoTenPT.Text = tbSVPT.Rows[0]["HOTEN"].ToString();
                }
            }
            else
            {
                cmbSachPT.SelectedIndex = -1;
                txtHoTenPT.Text = "";
            }
        }

        float giabia;
        private void cmbSachPT_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSachPT.SelectedIndex != -1)
            {
                DataTable tbPhieuMuon = dtBase.DataSelect("select * from PHIEUMUON where PHIEUMUON.MASACH = '" + cmbSachPT.SelectedValue.ToString() + "' and PHIEUMUON.MASV = '" + txtMSVPT.Text + "'");
                if (tbPhieuMuon.Rows.Count > 0)
                {
                    txtMaPhieuMuon.Text = tbPhieuMuon.Rows[0]["MAPHIEUMUON"].ToString();

                    DataTable tbSach = dtBase.DataSelect("select * from SACH join PHIEUMUON ON SACH.MASACH = PHIEUMUON.MASACH where PHIEUMUON.MASACH = '" + cmbSachPT.SelectedValue.ToString() + "'");
                    if (tbSach.Rows.Count > 0)
                    {
                        giabia = float.Parse(tbSach.Rows[0]["GIABIA"].ToString());
                    }
                    /////////////
                    if (tbPhieuMuon.Rows[0]["TRANGTHAI"].ToString() == "Đang mượn sách")
                    {
                        cmbViPham.DataSource = dtBase.DataSelect("select MAVIPHAM, ND_VIPHAM from VIPHAM where ND_VIPHAM like N'%Không vi phạm%' or ND_VIPHAM like N'%Trả đúng hạn%' ");

                    }
                    else if (tbPhieuMuon.Rows[0]["TRANGTHAI"].ToString() == "Quá hạn")
                    {
                        cmbViPham.DataSource = dtBase.DataSelect("select MAVIPHAM, ND_VIPHAM from VIPHAM where ND_VIPHAM like N'%Trả quá hạn%' ");

                    }
                    else if (tbPhieuMuon.Rows[0]["TRANGTHAI"].ToString() == "Không trả sách")
                    {
                        cmbViPham.DataSource = dtBase.DataSelect("select MAVIPHAM, ND_VIPHAM from VIPHAM where ND_VIPHAM like N'%Không trả sách%' ");

                    }
                    cmbViPham.DisplayMember = "ND_VIPHAM";
                    cmbViPham.ValueMember = "MAVIPHAM";
                }
            }
            else
            {
                txtMaPhieuMuon.Text = "";
            }
        }

        private void cmbNamPT_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvPhieuTra.DataSource = dtBase.DataSelect("Select PHIEUTRA.MAPHIEUTRA, PHIEUTRA.MAPHIEUMUON, SinhVien.MASV, SinhVien.HOTEN, SACH.TENSACH, PHIEUTRA.NGAYTRA, VIPHAM.ND_VIPHAM, VIPHAM.ND_PHAT, PHINOPPHAT, GHICHU " +
                "from ((((PHIEUTRA LEFT JOIN VIPHAM ON PHIEUTRA.MAVIPHAM = VIPHAM.MAVIPHAM) join PHIEUMUON on PHIEUMUON.MAPHIEUMUON = PHIEUTRA.MAPHIEUMUON) join SinhVien on PHIEUMUON.MASV = SinhVien.MaSV) join SACH ON SACH.MASACH = PHIEUMUON.MASACH) " +
                "where month(PHIEUTRA.NGAYTRA)= '" + cmbThangPT.Text + "' and  year(PHIEUTRA.NGAYTRA)= '" + cmbNamPT.Text + "'");
        }

        private void cmbLocPT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cmbLoc = "Select PHIEUTRA.MAPHIEUTRA, PHIEUTRA.MAPHIEUMUON, SinhVien.MASV, SinhVien.HOTEN, SACH.TENSACH, PHIEUTRA.NGAYTRA, VIPHAM.ND_VIPHAM, VIPHAM.ND_PHAT, PHINOPPHAT, GHICHU " +
                "from ((((PHIEUTRA LEFT JOIN VIPHAM ON PHIEUTRA.MAVIPHAM = VIPHAM.MAVIPHAM) join PHIEUMUON on PHIEUMUON.MAPHIEUMUON = PHIEUTRA.MAPHIEUMUON) join SinhVien on PHIEUMUON.MASV = SinhVien.MaSV) join Sach on SACH.MASACH = PHIEUMUON.MASACH) ";
            
            if (cmbLocPT.Text == "Tất cả")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc);
            }
            else if (cmbLocPT.Text == "Không vi phạm")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + " where VIPHAM.ND_VIPHAM = N'Không vi phạm'");
            }
            else if (cmbLocPT.Text == "Trả đúng hạn, Làm hỏng sách")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Trả đúng hạn, Làm hỏng sách'");
            }
            else if (cmbLocPT.Text == "Trả đúng hạn, Làm mất sách")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Trả đúng hạn, Làm mất sách'");
            }
            else if (cmbLocPT.Text == "Trả quá hạn")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Trả quá hạn'");
            }
            else if (cmbLocPT.Text == "Trả quá hạn, Làm hỏng sách")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Trả quá hạn, Làm hỏng sách'");
            }
            else if (cmbLocPT.Text == "Trả quá hạn, Làm mất sách")
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Trả quá hạn, Làm mất sách'");
            }
            else
            {
                dgvPhieuTra.DataSource = dtBase.DataSelect(cmbLoc + "where VIPHAM.ND_VIPHAM = N'Không trả sách'");
            }
        }

        private void cmbViPham_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbViPham.Text == "Trả đúng hạn, Làm hỏng sách")
            {
                txtPhiPhat.Text = (giabia * 20 / 100).ToString();
                cmbGhiChu.Enabled = true;
            }
            else if (cmbViPham.Text == "Trả đúng hạn, Làm mất sách")
            {
                txtPhiPhat.Text = (giabia * 50 / 100).ToString();
                cmbGhiChu.Enabled = true;
            }
            else if (cmbViPham.Text == "Trả quá hạn, Làm hỏng sách")
            {
                txtPhiPhat.Text = (giabia * 25 / 100).ToString();
                cmbGhiChu.Enabled = true;
            }
            else if (cmbViPham.Text == "Trả quá hạn, Làm mất sách")
            {
                txtPhiPhat.Text = (giabia * 55 / 100).ToString();
                cmbGhiChu.Enabled = true;
            }
            else if (cmbViPham.Text == "Không trả sách")
            {
                txtPhiPhat.Text = (giabia * 70 / 100).ToString();
                cmbGhiChu.Enabled = true;
            }
            else if (cmbViPham.Text == "Trả quá hạn" || cmbViPham.Text == "Không vi phạm")
            {
                txtPhiPhat.Text = "0";
                cmbGhiChu.Enabled = false;
            }
            else
            {
                txtPhiPhat.Text = "";
                cmbGhiChu.Enabled = true;
            }
        }
        public void ResetValuePTra()
        {
            txtMaPhieuTra.Text = "";
            txtMSVPT.Text = "";
            //cmbSachPT.SelectedIndex = -1;
            cmbSachPT.Text = "";
            txtMaPhieuMuon.Text = "";
            dtpNgayTra.Text = now.ToString();
            cmbViPham.SelectedIndex = -1;
            cmbViPham.Text = "";
            txtPhiPhat.Text = "";
            cmbGhiChu.Text = "";
            cmbThangPT.Text = now.Month.ToString();
            cmbNamPT.Text = now.Year.ToString();
            cmbLocPT.Text = "";
            giabia = 0;
        }

        private void reloadPT()
        {
            txtMaPhieuTra.Enabled = true;
            txtMSVPT.Enabled = true;
            dtpNgayTra.Enabled = true;
            btnTaoPhieuTra.Enabled = true;
            btnLuuPT.Enabled = false;
            btnUpPT.Enabled = false;
            btnSuaPhieuTra.Enabled = true;
            btnXoaPhieuTra.Enabled = true;

            btnTaoPhieuTra.BackColor = Color.Peru;
            btnLuuPT.BackColor = Color.PeachPuff;
            btnUpPT.BackColor = Color.PeachPuff;
            btnSuaPhieuTra.BackColor = Color.Peru;
            btnXoaPhieuTra.BackColor = Color.Peru;
        }
        private void btnThemPhieuTra_Click(object sender, EventArgs e)
        {
            txtTKPT.Text = "";
            ResetValuePTra();
            LoadDataPhieuTra();
            reloadPT();
        }
        private void btnTaoPhieuTra_Click(object sender, EventArgs e)
        {
            txtMaPhieuTra.Text = sinhma.SinhMa("PHIEUTRA", "PT", "MAPHIEUTRA");
            txtMaPhieuTra.Enabled = false;

            txtMSVPT.Text = "";
            cmbSachPT.Text = "";
            txtMaPhieuMuon.Text = "";
            dtpNgayTra.Text = now.ToString();
            cmbViPham.SelectedIndex = -1;
            cmbViPham.Text = "";
            txtPhiPhat.Text = "";
            cmbGhiChu.Text = "";
            giabia = 0;

            cmbViPham.Enabled = true;
            btnLuuPT.Enabled = true;
            btnUpPT.Enabled = false;
            btnSuaPhieuTra.Enabled = false;
            btnXoaPhieuTra.Enabled = false;

            btnLuuPT.BackColor = Color.Peru;
            btnUpPT.BackColor = Color.PeachPuff;
            btnSuaPhieuTra.BackColor = Color.PeachPuff;
            btnXoaPhieuTra.BackColor = Color.PeachPuff;

        }

        private void btnLuuPT_Click(object sender, EventArgs e)
        {
            if (txtMaPhieuTra.Text.Trim() == "" || txtMSVPT.Text.Trim() == "" || cmbSachPT.SelectedIndex == -1 || cmbViPham.SelectedIndex == -1 || cmbSachPT.Text.Trim() == "" || cmbViPham.Text.Trim()=="")
            {
                MessageBox.Show("Bạn chưa nhập đủ thông tin!\nVui lòng nhập đầy đủ thông tin để tiến hành tạo phiếu! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPhieuTra.Focus();
                return;
            }
            if (dtpNgayTra.Value.ToString().Substring(0, 11) != now.ToString().Substring(0, 11))
            {
                MessageBox.Show("Ngày tạo phiếu trả phải là ngày hôm nay!\n Bạn không được thay đổi ngày mượn sách!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dtpNgayTra.Focus();
                return;
            }
            if(cmbGhiChu.Enabled == true && cmbGhiChu.Text.Trim() == "")
            {
                MessageBox.Show("Vui lòng chọn phần ghi chú!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbGhiChu.Focus();
                return;
            }
            else
            {
                DataTable tbSV_ViPham = dtBase.DataSelect("select * from TAIKHOAN_BIKHOA where MASV_BIKHOA = '" + txtMSVPT.Text + "'");

                if (tbSV_ViPham.Rows.Count == 0 && txtMSVPT.Text != "")
                {
                    if (cmbViPham.Text == "Không trả sách")
                    {
                        string TaiKhoanBiKhoa = "insert into TAIKHOAN_BIKHOA values(N'" + txtMSVPT.Text + "')";
                        dtBase.Updatedate(TaiKhoanBiKhoa);
                    }
                }

                string sqlInsert = "insert into PHIEUTRA values(N'" + txtMaPhieuTra.Text + "', N'" + txtMaPhieuMuon.Text + "', N'" + dtpNgayTra.Value.ToString("yyyy-MM-dd") + "', N'" + cmbViPham.SelectedValue.ToString() + "', '" + float.Parse(txtPhiPhat.Text) + "', N'" + cmbGhiChu.Text + "')";
                dtBase.Updatedate(sqlInsert);

                DataTable tbPM = dtBase.DataSelect("select * from PHIEUMUON where MAPHIEUMUON = '" + txtMaPhieuMuon.Text + "'");
                if (tbPM.Rows[0]["TRANGTHAI"].ToString() == "Không trả sách")
                {
                    string sqlUPdateTT_KTra = "update PHIEUMUON set PHIEUMUON.TRANGTHAI = N'Không trả sách, nhân viên đã lập phiếu trả' WHERE PHIEUMUON.MAPHIEUMUON = '" + txtMaPhieuMuon.Text + "'";
                    dtBase.Updatedate(sqlUPdateTT_KTra);
                    LoadDataPhieuMuon();
                }
                else
                {
                    string sqlUPdateTrangThai = "update PHIEUMUON set PHIEUMUON.TRANGTHAI = N'Đã trả' WHERE PHIEUMUON.MAPHIEUMUON = '" + txtMaPhieuMuon.Text + "'";
                    dtBase.Updatedate(sqlUPdateTrangThai);
                    LoadDataPhieuMuon();
                }

                DataTable tbSach_PT = dtBase.DataSelect("select * from SACH join PHIEUMUON ON SACH.MASACH = PHIEUMUON.MASACH where PHIEUMUON.MASACH = '" + cmbSachPT.SelectedValue.ToString() + "'");
                int TongSach;
                if (tbSach_PT.Rows.Count > 0 && (cmbViPham.Text == "Không vi phạm" || cmbViPham.Text == "Trả đúng hạn, Làm hỏng sách" || cmbViPham.Text == "Trả quá hạn" || cmbViPham.Text == "Trả quá hạn, Làm hỏng sách"))
                {
                    TongSach = int.Parse(tbSach_PT.Rows[0]["TongSoLuong"].ToString());

                    TongSach = TongSach + 1;
                    string sqlUpdateSLSach = "update SACH set SACH.TONGSOLUONG = '" + TongSach.ToString() + "' where SACH.MASACH = '" + cmbSachPT.SelectedValue.ToString() + "'";
                    dtBase.Updatedate(sqlUpdateSLSach);
                    LoadDataSach();
                }

                MessageBox.Show("Bạn đã thêm thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetValuePTra();
                LoadDataPhieuTra();
                TinhTongPT();
                TinhTongPM();
                reloadPT();
            }
        }

        private void btnSuaPhieuTra_Click(object sender, EventArgs e)
        {
            txtMSVPT.Enabled = false;
            dtpNgayTra.Enabled = false;
            cmbViPham.Enabled = false;

            btnTaoPhieuTra.Enabled = false;
            btnLuuPT.Enabled = false;
            btnXoaPhieuTra.Enabled = false;
            btnUpPT.Enabled = true;

            btnTaoPhieuTra.BackColor = Color.PeachPuff;
            btnLuuPT.BackColor = Color.PeachPuff;
            btnXoaPhieuTra.BackColor = Color.PeachPuff;
            btnUpPT.BackColor = Color.Peru;

        }
        private void btnUpPT_Click(object sender, EventArgs e)
        {
            if (txtMaPhieuTra.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã phiếu trả!\nVui lòng nhập mã phiếu trả để tiến hành sửa! ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPhieuTra.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from PHIEUTRA where MAPHIEUTRA='" + txtMaPhieuTra.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaPhieuTra.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaPhieuTra.Focus();
                    ResetValuePTra();
                    return;
                }
                else
                {
                    if(tbDanhSach.Rows[0]["GHICHU"].ToString() == "Đã thanh toán phí nộp phạt")
                    {
                        MessageBox.Show("Sinh viên đã thanh toán phí nộp phạt!\nBạn không được sửa phần ghi chú này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        cmbGhiChu.Text = tbDanhSach.Rows[0]["GHICHU"].ToString();
                        return;
                    }
                    if (MessageBox.Show("Bạn có muốn sửa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        if (cmbGhiChu.SelectedIndex != -1)
                        {
                            string sqlUpdate = "update PHIEUTRA set GHICHU = (N'" + cmbGhiChu.Text + "') where MAPHIEUTRA='" + txtMaPhieuTra.Text + "'";
                            dtBase.Updatedate(sqlUpdate);
                        }
                        LoadDataPhieuTra();
                        reloadPT();
                    }
                }
            }
        }

        private void btnXoaPhieuTra_Click(object sender, EventArgs e)
        {
            if (txtMaPhieuTra.Text.Trim() == "")
            {
                MessageBox.Show("Bạn chưa nhập mã phiếu trả.\nVui lòng nhập mã phiếu trả để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPhieuTra.Focus();
                return;
            }
            else
            {
                DataTable tbDanhSach = dtBase.DataSelect("select * from PHIEUTRA where MAPHIEUTRA='" + txtMaPhieuTra.Text + "'");
                if (tbDanhSach.Rows.Count == 0)
                {
                    MessageBox.Show("Danh sách không có mã " + txtMaPhieuTra.Text + " \nVui lòng nhập mã khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtMaPhieuTra.Focus();
                    ResetValuePTra();
                    return;
                }

                if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string sqlDeletePM = "delete from PHIEUTRA where MAPHIEUTRA = (N'" + txtMaPhieuTra.Text + "')";
                    dtBase.Updatedate(sqlDeletePM);
                    LoadDataPhieuTra();
                    cmbViPham.SelectedIndex = -1;
                    ResetValuePTra();
                    TinhTongPT();
                }
            }
        }

        private void txtTKPT_TextChanged(object sender, EventArgs e)
        {
            dgvPhieuTra.DataSource = dtBase.DataSelect("Select PHIEUTRA.MAPHIEUTRA, PHIEUTRA.MAPHIEUMUON, SinhVien.MASV, SinhVien.HOTEN, SACH.TENSACH, PHIEUTRA.NGAYTRA, VIPHAM.ND_VIPHAM, VIPHAM.ND_PHAT, PHINOPPHAT, GHICHU " +
                "from ((((PHIEUTRA LEFT JOIN VIPHAM ON PHIEUTRA.MAVIPHAM = VIPHAM.MAVIPHAM) join PHIEUMUON on PHIEUMUON.MAPHIEUMUON = PHIEUTRA.MAPHIEUMUON) join SinhVien on PHIEUMUON.MASV = SinhVien.MaSV) join Sach on SACH.MASACH = PHIEUMUON.MASACH) " +
                "where MAPHIEUTRA like (N'%" + txtTKPT.Text + "%') or PHIEUMUON.MASV like (N'%" + txtTKPT.Text + "%')");
        }

        ///------------CellClick------------////
        private void dgvSinhVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaSV_SV.Text = dgvSinhVien.CurrentRow.Cells[0].Value.ToString();
            txtHoTen.Text = dgvSinhVien.CurrentRow.Cells[1].Value.ToString();
            dtpNgaySinh.Text = dgvSinhVien.CurrentRow.Cells[2].Value.ToString();
            cmbGioiTinh.Text = dgvSinhVien.CurrentRow.Cells[3].Value.ToString();
            txtDiaChi.Text = dgvSinhVien.CurrentRow.Cells[4].Value.ToString();
            txtSDT.Text = dgvSinhVien.CurrentRow.Cells[5].Value.ToString();
            txtEmail.Text = dgvSinhVien.CurrentRow.Cells[6].Value.ToString();
            
        }

        private void dgvSach_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaSach.Text = dgvSach.CurrentRow.Cells[0].Value.ToString();
            txtTenSach.Text = dgvSach.CurrentRow.Cells[1].Value.ToString();
            cmbTacGia.Text = dgvSach.CurrentRow.Cells[2].Value.ToString();
            cmbNXB.Text = dgvSach.CurrentRow.Cells[3].Value.ToString();
            cmbTheLoai.Text = dgvSach.CurrentRow.Cells[4].Value.ToString();
            txtSLSach.Text = dgvSach.CurrentRow.Cells[5].Value.ToString();
            txtGiaBia.Text = dgvSach.CurrentRow.Cells[6].Value.ToString();
            btnLuuSach.Enabled = false;
            btnLuuSach.BackColor = Color.PeachPuff;

        }

        private void dgvTacGia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaTacGia.Text = dgvTacGia.CurrentRow.Cells[0].Value.ToString();
            txtTenTacGia.Text = dgvTacGia.CurrentRow.Cells[1].Value.ToString();
            btnLuuTG.Enabled = false;
            btnLuuTG.BackColor = Color.PeachPuff;
        }
        private void dgvTheLoai_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaTheLoai.Text = dgvTheLoai.CurrentRow.Cells[0].Value.ToString();
            txtTenTheLoai.Text = dgvTheLoai.CurrentRow.Cells[1].Value.ToString();
            btnLuuTL.Enabled = false;
            btnLuuTL.BackColor = Color.PeachPuff;
        }

        private void dgvNXB_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaNXBan.Text = dgvNXB.CurrentRow.Cells[0].Value.ToString();
            txtTenNXB.Text = dgvNXB.CurrentRow.Cells[1].Value.ToString();
            txtDiaChiNXB.Text = dgvNXB.CurrentRow.Cells[2].Value.ToString();
            txtSDT_NXB.Text = dgvNXB.CurrentRow.Cells[3].Value.ToString();

            btnLuuNXB.Enabled = false;
            btnLuuNXB.BackColor = Color.PeachPuff;
        }

        private void dgvPhieuMuon_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaPM.Text = dgvPhieuMuon.CurrentRow.Cells[0].Value.ToString();
            txtMSV.Text = dgvPhieuMuon.CurrentRow.Cells[1].Value.ToString();
            cmbSach.Text = dgvPhieuMuon.CurrentRow.Cells[3].Value.ToString();
            txtSLMuon.Text = dgvPhieuMuon.CurrentRow.Cells[4].Value.ToString();
            dtpNgayMuon.Text = dgvPhieuMuon.CurrentRow.Cells[5].Value.ToString();
            dtpNgayHenTra.Text = dgvPhieuMuon.CurrentRow.Cells[6].Value.ToString();
            cmbTrangThai.Text = dgvPhieuMuon.CurrentRow.Cells[7].Value.ToString();
            btnLuuPM.Enabled = false;
            btnLuuPM.BackColor = Color.PeachPuff;
        }

        private void dgvPhieuTra_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaPhieuTra.Text = dgvPhieuTra.CurrentRow.Cells[0].Value.ToString();
            txtMaPhieuMuon.Text = dgvPhieuTra.CurrentRow.Cells[1].Value.ToString();
            txtMSVPT.Text = dgvPhieuTra.CurrentRow.Cells[2].Value.ToString();
            cmbSachPT.Text = dgvPhieuTra.CurrentRow.Cells[4].Value.ToString();
            dtpNgayTra.Text = dgvPhieuTra.CurrentRow.Cells[5].Value.ToString();
            cmbViPham.Text = dgvPhieuTra.CurrentRow.Cells[6].Value.ToString();
            txtPhiPhat.Text = dgvPhieuTra.CurrentRow.Cells[8].Value.ToString();
            cmbGhiChu.Text = dgvPhieuTra.CurrentRow.Cells[9].Value.ToString();
            btnLuuPT.Enabled = false;
            btnLuuPT.BackColor = Color.PeachPuff;
            if(cmbViPham.Text == "Không vi phạm" || cmbViPham.Text == "Trả quá hạn")
            {
                cmbGhiChu.Enabled = false;
            }
            else
            {
                cmbGhiChu.Enabled = true;
            }
        }

        private void btnXuatPM_Click(object sender, EventArgs e)
        {
            if (txtMaPM.Text.Trim() == "")
            {
                MessageBox.Show("Mời bạn nhập vào mã phiếu mượn để tiến hành lưu file excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataTable tbP_Muon = dtBase.DataSelect("select * from (PHIEUMUON join SACH on PHIEUMUON.MASACH = SACH.MASACH) join SINHVIEN ON SINHVIEN.MASV = PHIEUMUON.MASV where PHIEUMUON.MAPHIEUMUON = '" + txtMaPM.Text + "'");
            if (tbP_Muon.Rows.Count == 0)
            {
                MessageBox.Show("Không tồn tại phiếu mượn này!\nBạn không thể xuất file excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPM.Focus();
                ResetvaluePM();
                return;
            }
            if(tbP_Muon.Rows[0]["TRANGTHAI"].ToString() != "Đang mượn sách")
            {
                MessageBox.Show("Bạn không được xuất phiếu mượn này! Bạn chỉ được xuất những phiếu mượn có trạng thái 'Đang mượn sách'", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPM.Focus();
                ResetvaluePM();
                return;
            }
            if (tbP_Muon.Rows.Count > 0)
            {
                //Khai báo và khởi tạo các đối tượng
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];

                exSheet.get_Range("A1:A14").ColumnWidth = 20;
                exSheet.get_Range("B1:B14").ColumnWidth = 35;
                exSheet.get_Range("A1:B16").Font.Name = "Times new roman";

                //Định dạng chung
                Excel.Range DHGTVT = (Excel.Range)exSheet.Cells[1, 1];
                exSheet.get_Range("A1:B1").Merge(true);
                DHGTVT.Font.Size = 13;
                DHGTVT.Font.Bold = true;
                DHGTVT.Value = "TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI";
                exSheet.get_Range("A1:B1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////
                Excel.Range PhieuMuon = (Excel.Range)exSheet.Cells[3, 1];
                exSheet.get_Range("A3:B3").Merge(true);
                PhieuMuon.Font.Size = 12;
                PhieuMuon.Font.Bold = true;
                PhieuMuon.Value = "PHIẾU MƯỢN SÁCH";
                exSheet.get_Range("A3:B3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                exSheet.get_Range("A5").Value = "Mã phiếu mượn: ";
                exSheet.get_Range("A6").Value = "Mã sinh viên: ";
                exSheet.get_Range("A7").Value = "Tên sinh viên: ";
                exSheet.get_Range("A8").Value = "Tên sách mượn: ";
                exSheet.get_Range("A9").Value = "Tên sách mượn: ";
                exSheet.get_Range("A10").Value = "Ngày mượn: ";
                exSheet.get_Range("A11").Value = "Ngày hẹn trả: ";
                exSheet.get_Range("A12").Value = "Trạng thái mượn: ";


                exSheet.get_Range("B5").Value = tbP_Muon.Rows[0]["MAPHIEUMUON"].ToString();
                exSheet.get_Range("B6").Value = tbP_Muon.Rows[0]["MASV"].ToString();
                exSheet.get_Range("B7").Value = tbP_Muon.Rows[0]["HOTEN"].ToString();
                exSheet.get_Range("B8").Value = tbP_Muon.Rows[0]["TENSACH"].ToString();
                exSheet.get_Range("B9").Value = tbP_Muon.Rows[0]["SOLUONGMUON"].ToString();
                exSheet.get_Range("B10").Value = tbP_Muon.Rows[0]["NGAYMUON"].ToString();
                exSheet.get_Range("B11").Value = tbP_Muon.Rows[0]["NGAYHENTRA"].ToString();
                exSheet.get_Range("B12").Value = tbP_Muon.Rows[0]["TRANGTHAI"].ToString();

                exSheet.get_Range("A5:B12").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                exSheet.get_Range("B14:B16").Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                exSheet.get_Range("B14").Value = "Hà Nội, ngày " + now.Day.ToString() + ", tháng " + now.Month.ToString() + ", năm " + now.Year.ToString();
                exSheet.get_Range("B15").Value = "Ký tên và đóng dấu xác nhận";
                exSheet.get_Range("B16").Value = "Ghi rõ họ tên ";

                exSheet.Name = "PhieuMuon"; //dat ten cho sheet dang lam viec
                exBook.Activate(); //Kích hoạt file Excel

                //Thiết lập các thuộc tính của SaveFileDialog
                SaveFileDialog dlgSave = new SaveFileDialog();
                dlgSave.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.docx) | *.docx | All files(*.*) | *.* ";
                dlgSave.FilterIndex = 1;
                dlgSave.AddExtension = true;
                dlgSave.DefaultExt = ".xlsx";
                if (dlgSave.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    exBook.SaveAs(dlgSave.FileName.ToString());//Lưu file Excel
                    MessageBox.Show("Lưu file thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                exApp.Quit();//Thoát khỏi ứng dụng
            }
        }

        private void btnXuatPT_Click(object sender, EventArgs e)
        {
            if (txtMaPhieuTra.Text.Trim() == "")
            {
                MessageBox.Show("Mời bạn nhập vào mã phiếu trả để tiến hành xuất file excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataTable tbP_Tra = dtBase.DataSelect("select * from ((((PHIEUTRA LEFT JOIN VIPHAM ON PHIEUTRA.MAVIPHAM = VIPHAM.MAVIPHAM) " +
                "join PHIEUMUON on PHIEUMUON.MAPHIEUMUON = PHIEUTRA.MAPHIEUMUON) join SinhVien on PHIEUMUON.MASV = SinhVien.MaSV) " +
                "join Sach on SACH.MASACH = PHIEUMUON.MASACH) where MAPHIEUTRA = '" + txtMaPhieuTra.Text + "'");
            if (tbP_Tra.Rows.Count == 0)
            {
                MessageBox.Show("Không tồn tại phiếu trả này!\nBạn không thể xuất file excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMaPhieuTra.Focus();
                ResetValuePTra();
                return;
            }
            if (tbP_Tra.Rows[0]["ND_VIPHAM"].ToString() == "Không trả sách")
            {
                MessageBox.Show("Bạn không thể xuất phiếu trả cho phiếu mượn này!\nLý do: Nội dung vi pham là: không trả sách!", "Thông báo");
                txtMaPhieuTra.Focus();
                ResetValuePTra();
                return;
            }
            
            if (tbP_Tra.Rows.Count > 0)
            {
                //Khai báo và khởi tạo các đối tượng
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];

                exSheet.get_Range("A1:A17").ColumnWidth = 20;
                exSheet.get_Range("B1:B17").ColumnWidth = 35;
                exSheet.get_Range("A1:B17").Font.Name = "Times new roman";

                //Định dạng chung
                Excel.Range DHGTVT = (Excel.Range)exSheet.Cells[1, 1];
                exSheet.get_Range("A1:B1").Merge(true);
                DHGTVT.Font.Size = 13;
                DHGTVT.Font.Bold = true;
                DHGTVT.Value = "TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI";
                exSheet.get_Range("A1:B1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////
                Excel.Range PhieuTra = (Excel.Range)exSheet.Cells[3, 1];
                exSheet.get_Range("A3:B3").Merge(true);
                PhieuTra.Font.Size = 12;
                PhieuTra.Font.Bold = true;
                PhieuTra.Value = "PHIẾU TRẢ SÁCH";
                exSheet.get_Range("A3:B3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                exSheet.get_Range("A5").Value = "Mã phiếu trả: ";
                exSheet.get_Range("A6").Value = "Mã phiếu mượn: ";
                exSheet.get_Range("A7").Value = "Mã sinh viên: ";
                exSheet.get_Range("A8").Value = "Tên sinh viên: ";
                exSheet.get_Range("A9").Value = "Tên sách trả: ";
                exSheet.get_Range("A10").Value = "Ngày trả sách: ";
                exSheet.get_Range("A11").Value = "Vi phạm: ";
                exSheet.get_Range("A12").Value = "Phí nộp phạt: ";
                exSheet.get_Range("A13").Value = "Ghi chú: ";


                exSheet.get_Range("B5").Value = tbP_Tra.Rows[0]["MAPHIEUTRA"].ToString();
                exSheet.get_Range("B6").Value = tbP_Tra.Rows[0]["MAPHIEUMUON"].ToString();
                exSheet.get_Range("B7").Value = tbP_Tra.Rows[0]["MASV"].ToString();
                exSheet.get_Range("B8").Value = tbP_Tra.Rows[0]["HOTEN"].ToString();
                exSheet.get_Range("B9").Value = tbP_Tra.Rows[0]["TENSACH"].ToString();
                exSheet.get_Range("B10").Value = tbP_Tra.Rows[0]["NGAYTRA"].ToString();
                exSheet.get_Range("B11").Value = tbP_Tra.Rows[0]["ND_VIPHAM"].ToString();
                exSheet.get_Range("B12").Value = tbP_Tra.Rows[0]["PHINOPPHAT"].ToString();
                exSheet.get_Range("B13").Value = tbP_Tra.Rows[0]["GHICHU"].ToString();

                exSheet.get_Range("A5:B13").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                exSheet.get_Range("B15:B17").Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                exSheet.get_Range("B15").Value = "Hà Nội, ngày " + now.Day.ToString() + ", tháng " + now.Month.ToString() + ", năm " + now.Year.ToString();
                exSheet.get_Range("B16").Value = "Ký tên và đóng dấu xác nhận";
                exSheet.get_Range("B17").Value = "Ghi rõ họ tên ";


                exSheet.Name = "PhieuTra"; //dat ten cho sheet dang lam viec
                exBook.Activate(); //Kích hoạt file Excel

                //Thiết lập các thuộc tính của SaveFileDialog
                SaveFileDialog dlgSave = new SaveFileDialog();
                dlgSave.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.docx) | *.docx | All files(*.*) | *.* ";
                dlgSave.FilterIndex = 1;
                dlgSave.AddExtension = true;
                dlgSave.DefaultExt = ".xlsx";
                if (dlgSave.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    exBook.SaveAs(dlgSave.FileName.ToString());//Lưu file Excel
                    MessageBox.Show("Lưu file thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                exApp.Quit();//Thoát khỏi ứng dụng
            }
        }

        private void btnInSV_Click(object sender, EventArgs e)
        {
            DataTable tbSV = dtBase.DataSelect("select * from SINHVIEN");

            if (tbSV.Rows.Count > 0)
            {
                //Khai báo và khởi tạo các đối tượng
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];

                exSheet.get_Range("A1").ColumnWidth = 7;
                exSheet.get_Range("B1").ColumnWidth = 15;
                exSheet.get_Range("C1").ColumnWidth = 25;
                exSheet.get_Range("D1").ColumnWidth = 20;
                exSheet.get_Range("E1").ColumnWidth = 8;
                exSheet.get_Range("F1").ColumnWidth = 25;
                exSheet.get_Range("G1").ColumnWidth = 12;
                exSheet.get_Range("H1").ColumnWidth = 30;

                //Định dạng chung
                Excel.Range DHGTVT = (Excel.Range)exSheet.Cells[2, 1];
                exSheet.get_Range("A2:C2").Merge(true);
                DHGTVT.Font.Size = 13;
                DHGTVT.Font.Bold = true;
                DHGTVT.Value = "TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI";

                ////
                Excel.Range CHXHCNVN = (Excel.Range)exSheet.Cells[2, 5];
                exSheet.get_Range("E2:H2").Merge(true);
                CHXHCNVN.Font.Size = 13;
                CHXHCNVN.Font.Bold = true;
                CHXHCNVN.Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";

                ////
                Excel.Range DLTDHP = (Excel.Range)exSheet.Cells[3, 5];
                exSheet.get_Range("E3:H3").Merge(true);
                DLTDHP.Font.Size = 11;
                DLTDHP.Value = "Độc Lập - Tự Do - Hạnh Phúc";

                ////
                Excel.Range DANHSACH = (Excel.Range)exSheet.Cells[6, 1];
                exSheet.get_Range("A6:H6").Merge(true);
                DANHSACH.Font.Size = 12;
                DANHSACH.Font.Bold = true;
                DANHSACH.Value = "DANH SÁCH SINH VIÊN";

                ///
                exSheet.get_Range("A8:H8").Font.Bold = true;
                exSheet.get_Range("A8").Value = "STT";
                exSheet.get_Range("B8").Value = "Mã sinh viên";
                exSheet.get_Range("C8").Value = "Họ Tên";
                exSheet.get_Range("D8").Value = "Ngày sinh";
                exSheet.get_Range("E8").Value = "Giới tính";
                exSheet.get_Range("F8").Value = "Địa chỉ";
                exSheet.get_Range("G8").Value = "SĐT";
                exSheet.get_Range("H8").Value = "Email";

                exSheet.get_Range("A2:H8").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////

                for (int i = 0; i < tbSV.Rows.Count; i++)
                {
                    exSheet.get_Range("A" + (i + 9).ToString() + ":G" + (i + 9).ToString()).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    exSheet.get_Range("A" + (i + 9).ToString() + ":G" + (i + 9).ToString()).Font.Bold = false;
                    exSheet.get_Range("A" + (i + 9).ToString()).Value = (i + 1).ToString();
                    exSheet.get_Range("B" + (i + 9).ToString()).Value = tbSV.Rows[i]["MASV"].ToString();
                    exSheet.get_Range("C" + (i + 9).ToString()).Value = tbSV.Rows[i]["HOTEN"].ToString();
                    exSheet.get_Range("D" + (i + 9).ToString()).Value = tbSV.Rows[i]["NGAYSINH"].ToString();
                    exSheet.get_Range("E" + (i + 9).ToString()).Value = tbSV.Rows[i]["GIOITINH"].ToString();
                    exSheet.get_Range("F" + (i + 9).ToString()).Value = tbSV.Rows[i]["DIAVHI"].ToString();
                    exSheet.get_Range("G" + (i + 9).ToString()).Value = tbSV.Rows[i]["SDT"].ToString();
                    exSheet.get_Range("H" + (i + 9).ToString()).Value = tbSV.Rows[i]["EMAIL"].ToString();

                }
                exSheet.Name = "DanhSachSinhVien";
                exBook.Activate(); //Kích hoạt file Excel

                //Thiết lập các thuộc tính của SaveFileDialog
                SaveFileDialog dlgSave = new SaveFileDialog();
                dlgSave.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.docx)| *.docx | All files(*.*) | *.* ";

                dlgSave.FilterIndex = 1;
                dlgSave.AddExtension = true; //tự động thêm vào đuôi mở rộng nếu ng dùng k chỉ ra
                dlgSave.DefaultExt = ".xlsX"; ///đuôi mở rộng mặc định
                if (dlgSave.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    exBook.SaveAs(dlgSave.FileName.ToString());//Lưu file Excel
                    MessageBox.Show("Lưu file thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                exApp.Quit();//Thoát khỏi ứng dụng
            }
            else
            {
                MessageBox.Show("Không có danh sách hàng để in", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
