using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

namespace BTL_ThuVien.Module
{
    class Function
    {
        DataAccess data = new DataAccess();
        public string SinhMa(string TenBang, string MaBatDau, string TruongSinhMa)
        {
            int id = 1;
            bool ktra = false;
            string ma = "";
            DataTable tbBang = new DataTable();
            while (ktra == false) {
                tbBang = data.DataSelect("select * from " + TenBang + " where " + TruongSinhMa + " = '" + MaBatDau + id.ToString() + "'");
                if (tbBang.Rows.Count == 0)
                {
                    ktra = true;
                }
                else
                {
                    id++;
                    ktra = false;
                }
            }
            ma = MaBatDau + id.ToString();
            return ma;
        }
    }
}
