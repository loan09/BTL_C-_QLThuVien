using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace BTL_ThuVien.Module
{
    class DataAccess
    {
        string connectString = "Data Source=LAPTOP-MBVJHROR\\SQLEXPRESS;Initial Catalog=BTL_QLThuVienLTTQ10;Integrated Security=True";

        SqlConnection sqlConnect = null;

        ///mở kết nối
        void OpenConnect()
        {
            sqlConnect = new SqlConnection(connectString);
            if (sqlConnect.State != ConnectionState.Open)
            {
                sqlConnect.Open();
            }
        }
        //phương thức đóng kết nối
        void closeConnect()
        {
            if (sqlConnect.State != ConnectionState.Closed)
            {
                sqlConnect.Close(); //pt đóng kết nối
            }
            sqlConnect.Dispose(); //phương thức hủy đối tượng
        }
        ///phương thức thực hiện câu lệnh select trả về datatable
        public DataTable DataSelect(string sqlSelect)
        {
            DataTable dtResult = new DataTable();
            OpenConnect();
            SqlDataAdapter sqlData = new SqlDataAdapter(sqlSelect, sqlConnect);
            sqlData.Fill(dtResult);
            closeConnect();
            sqlData.Dispose();
            return dtResult;
        }

        ///phương thức thực hiện thay đổi dữ liệu: insert, delete, update
        public void Updatedate(string sql)
        {
            OpenConnect();
            SqlCommand sqlcom = new SqlCommand();
            sqlcom.Connection = sqlConnect;
            sqlcom.CommandText = sql;
            sqlcom.ExecuteNonQuery();
            closeConnect();
            sqlcom.Dispose();
        }
    }
}
