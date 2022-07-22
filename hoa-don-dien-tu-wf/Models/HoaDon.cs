using System;
using System.Collections.Generic;
using System.Text;

namespace hoa_don_dien_tu_wf.Models
{
    public class HoaDon
    {
        public string khhdon { get; set; } // ki hieu hoa don
        public string nmten { get; set; }
        public string nmmst { get; set; }
        public string nmdchi { get; set; }
        public string shdon { get; set; }
        public DateTime? nky { get; set; } // ngay hop dong
        public string nbten { get; set; }
        public string nbmst { get; set; }
        public string nbdchi { get; set; }
        public float? TSTSauThue { get; set; }
        public float? TongVAT { get; set; }
        public string LoaiTien { get; set; }
        public float? TiGia { get; set; }

        // hddv
        public List<HDDichVu> HDDichVuList { get; set; }

        public string CheckUnique { get; set; }
    }
}
