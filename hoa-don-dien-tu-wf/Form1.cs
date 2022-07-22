using ClosedXML.Excel;
using hoa_don_dien_tu_wf.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hoa_don_dien_tu_wf
{
    public partial class Form1 : Form
    {
        private string _saveDirPath;
        private static string _excutingPath = Process.GetCurrentProcess().MainModule.FileName;
        private static string _excutingDir = Path.GetDirectoryName(_excutingPath);
        private string _companyName;

        public Form1()
        {
            InitializeComponent();
            richTextBox1.TextChanged += richTextBox1_TextChanged;

            _saveDirPath = Path.Combine(_excutingDir, "Result");
            textBox2.Text = _saveDirPath.ToString();

            // init date
            dateTimePicker1.Value = DateTime.Now.AddDays(-30);
            dateTimePicker2.Value = DateTime.Now;
            //
            _companyName = "";
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = _saveDirPath;
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _saveDirPath = folderBrowserDialog1.SelectedPath;
                textBox2.Text = _saveDirPath.ToString();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        // check bat dau
        private async void button2_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox1.Text = "";
                richTextBox1.Text += "Bắt đầu chạy!";

                //disable 
                button2.Enabled = false;
                button3.Enabled = false;
                // check token
                var token = textBox1.Text.Trim();
                if(string.IsNullOrEmpty(token))
                {
                    throw new Exception("Token đang bị trống!");
                }
                // check time
                var startDate = dateTimePicker1.Value;
                var endDate = dateTimePicker2.Value;
                if(startDate > endDate)
                {
                    throw new Exception("Ngày bắt đầu phải sớm hơn ngày kết thúc!");
                }
                var textStartDate = startDate.ToString("dd/MM/yyyy");
                var textEndDate = endDate.ToString("dd/MM/yyyy");

                richTextBox1.Text += "\n\nĐang lấy thông tin hóa đơn bán và mua...";
                // lay tong hoa don
                var urlban = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=1000000&search=tdlap=ge={textStartDate}T00:00:00;tdlap=le={textEndDate}T23:59:59";
                var urlmua = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=1000000&search=tdlap=ge={textStartDate}T00:00:00;tdlap=le={textEndDate}T23:59:59";

                var danhsachhoadonban = new List<HoaDon>();
                var danhsachhoadonmua = new List<HoaDon>();
                //Console.WriteLine($"Đang lấy danh sách hóa đơn bán và mua ...");

                var firstTryCount = 0;
                while (firstTryCount < 5)
                {
                    try
                    {
                        using (var http = new HttpClient())
                        {
                            http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                            http.Timeout = TimeSpan.FromSeconds(45);

                            // lay danh sach ban
                            var responseban = await http.GetAsync(urlban);
                            if (responseban.StatusCode != System.Net.HttpStatusCode.OK)
                            {
                                throw new Exception($"Mã lỗi: {responseban.StatusCode} - {responseban.Content.ToString()}");
                            }
                            var resultban = JsonConvert.DeserializeObject<dynamic>(await responseban.Content.ReadAsStringAsync());
                            var datasban = resultban.datas;
                            danhsachhoadonban = JsonConvert.DeserializeObject<List<HoaDon>>(JsonConvert.SerializeObject(datasban));

                            // lay danh sach mua
                            var responsemua = await http.GetAsync(urlmua);
                            if (responsemua.StatusCode != System.Net.HttpStatusCode.OK)
                            {
                                throw new Exception($"Mã lỗi: {responsemua.StatusCode} - {responsemua.Content.ToString()}");
                            }
                            var resultmua = JsonConvert.DeserializeObject<dynamic>(await responsemua.Content.ReadAsStringAsync());
                            var datasmua = resultmua.datas;
                            danhsachhoadonmua = JsonConvert.DeserializeObject<List<HoaDon>>(JsonConvert.SerializeObject(datasmua));
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        firstTryCount++;
                    }
                }

                if (firstTryCount >= 5)
                {
                    //Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                    //Console.ReadKey();
                    throw new Exception("Lỗi!!! Xin lấy lại token và chạy lại!");
                }

                // lấy tên hiển thị
                if(danhsachhoadonban.Count == 0 && danhsachhoadonmua.Count == 0)
                {
                    throw new Exception("Không có hóa đơn nào trong khoảng thời gian này!!!");
                }

                if(danhsachhoadonban.Count > 0)
                {
                    _companyName = danhsachhoadonban[0].nbten;
                }
                else
                {
                    _companyName = danhsachhoadonmua[0].nmten;
                }

                // set company name
                label4.Text = _companyName;

                richTextBox1.Text += "\nLấy thông tin hóa đơn bán và mua thành công!";
                richTextBox1.Text += $"\n\nCÓ TẤT CẢ {danhsachhoadonban.Count} HÓA ĐƠN BÁN và {danhsachhoadonmua.Count} HÓA ĐƠN MUA!";


                richTextBox1.Text += "\n\nĐang lấy thông tin chi tiết hóa đơn bán...";
                int lenthText = richTextBox1.Text.Length;
                int indexban = 0;
                foreach (var hoadon in danhsachhoadonban)
                {
                    indexban++;
                    richTextBox1.Text = richTextBox1.Text.Substring(0, lenthText);
                    richTextBox1.Text += $"\n{indexban} / {danhsachhoadonban.Count}";
                    var detailUrl = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={hoadon.nbmst}&khhdon={hoadon.khhdon}&shdon={hoadon.shdon}&khmshdon=1";
                    var secondTryCount = 0;
                    while (secondTryCount < 10)
                    {
                        try
                        {
                            using (var http = new HttpClient())
                            {
                                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                                http.Timeout = TimeSpan.FromSeconds(45);
                                var response = await http.GetAsync(detailUrl);

                                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                                {
                                    throw new Exception($"Mã lỗi: {response.StatusCode} - {response.Content.ToString()}");
                                }
                                var result = JsonConvert.DeserializeObject<dynamic>(await response.Content.ReadAsStringAsync());

                                hoadon.TSTSauThue = float.Parse(result?.tgtttbso.ToString());
                                hoadon.TongVAT = float.Parse(result?.tgtthue.ToString());
                                hoadon.LoaiTien = result?.dvtte;
                                hoadon.TiGia = result?.tgia;

                                hoadon.HDDichVuList = JsonConvert.DeserializeObject<List<HDDichVu>>(JsonConvert.SerializeObject(result?.hdhhdvu));
                                break;
                            }
                        }
                        catch (Exception exx)
                        {
                            secondTryCount++;
                        }
                    }

                    if (secondTryCount >= 10)
                    {
                        //Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                        //Console.ReadKey();
                        throw new Exception("Lỗi!!! Xin lấy lại token và chạy lại!");
                    }
                }
                richTextBox1.Text += "\nLấy thông tin chi tiết hóa đơn bán thành công";

                richTextBox1.Text += "\n\nĐang lấy thông tin chi tiết hóa đơn mua...";
                int indexmua = 0;
                lenthText = richTextBox1.Text.Length;
                foreach (var hoadon in danhsachhoadonmua)
                {
                    indexmua++;
                    richTextBox1.Text = richTextBox1.Text.Substring(0, lenthText);
                    richTextBox1.Text += $"\n{indexmua} / {danhsachhoadonmua.Count}";
                    var detailUrl = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={hoadon.nbmst}&khhdon={hoadon.khhdon}&shdon={hoadon.shdon}&khmshdon=1";
                    var secondTryCount = 0;
                    while (secondTryCount < 10)
                    {
                        try
                        {
                            using (var http = new HttpClient())
                            {
                                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                                http.Timeout = TimeSpan.FromSeconds(45);
                                var response = await http.GetAsync(detailUrl);

                                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                                {
                                    throw new Exception($"Mã lỗi: {response.StatusCode} - {response.Content.ToString()}");
                                }
                                var result = JsonConvert.DeserializeObject<dynamic>(await response.Content.ReadAsStringAsync());

                                hoadon.TSTSauThue = float.Parse(result?.tgtttbso.ToString());
                                hoadon.TongVAT = float.Parse(result?.tgtthue.ToString());
                                hoadon.LoaiTien = result?.dvtte;
                                hoadon.TiGia = result?.tgia;

                                hoadon.HDDichVuList = JsonConvert.DeserializeObject<List<HDDichVu>>(JsonConvert.SerializeObject(result?.hdhhdvu));
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            secondTryCount++;
                        }
                    }

                    if (secondTryCount >= 10)
                    {
                        //Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                        //Console.ReadKey();
                        throw new Exception("Lỗi!!! Xin lấy lại token và chạy lại!");
                    }
                }

                richTextBox1.Text += "\nLấy thông tin chi tiết hóa đơn mua thành công!";

                richTextBox1.Text += "\n\nĐang in ra file...";
                // ban
                var workbook = new XLWorkbook();
                IXLWorksheet workSheetBan = workbook.Worksheets.Add($"Hóa Đơn Bán");
                workSheetBan.Cell(1, 1).Value = "Tên người mua";
                workSheetBan.Cell(1, 2).Value = "MST người mua";
                workSheetBan.Cell(1, 3).Value = "Địa chỉ người mua";
                workSheetBan.Cell(1, 4).Value = "Số HĐ";
                workSheetBan.Cell(1, 5).Value = "Ngày HĐ";
                workSheetBan.Cell(1, 6).Value = "Tên người bán";
                workSheetBan.Cell(1, 7).Value = "MST Người bán";
                workSheetBan.Cell(1, 8).Value = "Địa chỉ người bán";
                workSheetBan.Cell(1, 9).Value = "Tổng số tiền sau thuế";
                workSheetBan.Cell(1, 10).Value = "Tổng VAT";
                workSheetBan.Cell(1, 11).Value = "Loại tiền";
                workSheetBan.Cell(1, 12).Value = "Tỷ giá";
                workSheetBan.Cell(1, 13).Value = "STT";
                workSheetBan.Cell(1, 14).Value = "Tên hàng hóa dịch vụ";
                workSheetBan.Cell(1, 15).Value = "Đơn vị tính";
                workSheetBan.Cell(1, 16).Value = "Số lượng";
                workSheetBan.Cell(1, 17).Value = "Đơn giá";
                workSheetBan.Cell(1, 18).Value = "Tỷ lệ CK";
                workSheetBan.Cell(1, 19).Value = "Số tiền Ck";
                workSheetBan.Cell(1, 20).Value = "Thành tiền";
                workSheetBan.Cell(1, 21).Value = "Thuế Suất";
                workSheetBan.Cell(1, 22).Value = "Tiền thuế VAT";
                workSheetBan.Cell(1, 23).Value = "Check Unique";

                var printIndex1 = 0;
                for (int i = 0; i < danhsachhoadonban.Count; i++)
                {
                    var row = danhsachhoadonban[i];
                    for (int j = 0; j < row.HDDichVuList.Count; j++)
                    {
                        var hddv = row.HDDichVuList[j];
                        workSheetBan.Cell(2 + printIndex1, 1).Value = row.nmten;
                        workSheetBan.Cell(2 + printIndex1, 2).Value = row.nmmst;
                        workSheetBan.Cell(2 + printIndex1, 3).Value = row.nmdchi;
                        workSheetBan.Cell(2 + printIndex1, 4).Value = row.shdon;
                        workSheetBan.Cell(2 + printIndex1, 5).Value = row.nky.ToString();
                        workSheetBan.Cell(2 + printIndex1, 6).Value = row.nbten;
                        workSheetBan.Cell(2 + printIndex1, 7).Value = row.nbmst;
                        workSheetBan.Cell(2 + printIndex1, 8).Value = row.nbdchi;
                        workSheetBan.Cell(2 + printIndex1, 9).Value = row.TSTSauThue.ToString();
                        workSheetBan.Cell(2 + printIndex1, 10).Value = row.TongVAT.ToString();
                        workSheetBan.Cell(2 + printIndex1, 11).Value = row.LoaiTien;
                        workSheetBan.Cell(2 + printIndex1, 12).Value = row.TiGia.ToString();
                        workSheetBan.Cell(2 + printIndex1, 13).Value = (j + 1).ToString();


                        workSheetBan.Cell(2 + printIndex1, 14).Value = hddv.ten;
                        workSheetBan.Cell(2 + printIndex1, 15).Value = hddv.dvtinh;
                        workSheetBan.Cell(2 + printIndex1, 16).Value = hddv.sluong.ToString();
                        workSheetBan.Cell(2 + printIndex1, 17).Value = hddv.dgia.ToString();
                        workSheetBan.Cell(2 + printIndex1, 18).Value = hddv.tlckhau == null ? "-" : hddv.tlckhau.ToString();
                        workSheetBan.Cell(2 + printIndex1, 19).Value = hddv.stckhau == null ? "-" : hddv.stckhau.ToString();
                        workSheetBan.Cell(2 + printIndex1, 20).Value = hddv.thtien.ToString();
                        workSheetBan.Cell(2 + printIndex1, 21).Value = hddv.tsuat.ToString();
                        workSheetBan.Cell(2 + printIndex1, 22).Value = hddv.tsuat == null ? "0" : (hddv.tsuat * hddv.thtien).ToString();
                        workSheetBan.Cell(2 + printIndex1, 23).Value = "";

                        // tang index
                        printIndex1++;
                    }
                }


                // mua
                IXLWorksheet workbookMua = workbook.Worksheets.Add($"Hóa Đơn Mua");
                workbookMua.Cell(1, 1).Value = "Tên người mua";
                workbookMua.Cell(1, 2).Value = "MST người mua";
                workbookMua.Cell(1, 3).Value = "Địa chỉ người mua";
                workbookMua.Cell(1, 4).Value = "Số HĐ";
                workbookMua.Cell(1, 5).Value = "Ngày HĐ";
                workbookMua.Cell(1, 6).Value = "Tên người bán";
                workbookMua.Cell(1, 7).Value = "MST Người bán";
                workbookMua.Cell(1, 8).Value = "Địa chỉ người bán";
                workbookMua.Cell(1, 9).Value = "Tổng số tiền sau thuế";
                workbookMua.Cell(1, 10).Value = "Tổng VAT";
                workbookMua.Cell(1, 11).Value = "Loại tiền";
                workbookMua.Cell(1, 12).Value = "Tỷ giá";
                workbookMua.Cell(1, 13).Value = "STT";
                workbookMua.Cell(1, 14).Value = "Tên hàng hóa dịch vụ";
                workbookMua.Cell(1, 15).Value = "Đơn vị tính";
                workbookMua.Cell(1, 16).Value = "Số lượng";
                workbookMua.Cell(1, 17).Value = "Đơn giá";
                workbookMua.Cell(1, 18).Value = "Tỷ lệ CK";
                workbookMua.Cell(1, 19).Value = "Số tiền Ck";
                workbookMua.Cell(1, 20).Value = "Thành tiền";
                workbookMua.Cell(1, 21).Value = "Thuế Suất";
                workbookMua.Cell(1, 22).Value = "Tiền thuế VAT";
                workbookMua.Cell(1, 23).Value = "Check Unique";

                var printIndex2 = 0;
                for (int i = 0; i < danhsachhoadonmua.Count; i++)
                {
                    var row = danhsachhoadonmua[i];
                    for (int j = 0; j < row.HDDichVuList.Count; j++)
                    {
                        var hddv = row.HDDichVuList[j];
                        workbookMua.Cell(2 + printIndex2, 1).Value = row.nmten;
                        workbookMua.Cell(2 + printIndex2, 2).Value = row.nmmst;
                        workbookMua.Cell(2 + printIndex2, 3).Value = row.nmdchi;
                        workbookMua.Cell(2 + printIndex2, 4).Value = row.shdon;
                        workbookMua.Cell(2 + printIndex2, 5).Value = row.nky.ToString();
                        workbookMua.Cell(2 + printIndex2, 6).Value = row.nbten;
                        workbookMua.Cell(2 + printIndex2, 7).Value = row.nbmst;
                        workbookMua.Cell(2 + printIndex2, 8).Value = row.nbdchi;
                        workbookMua.Cell(2 + printIndex2, 9).Value = row.TSTSauThue.ToString();
                        workbookMua.Cell(2 + printIndex2, 10).Value = row.TongVAT.ToString();
                        workbookMua.Cell(2 + printIndex2, 11).Value = row.LoaiTien;
                        workbookMua.Cell(2 + printIndex2, 12).Value = row.TiGia.ToString();
                        workbookMua.Cell(2 + printIndex2, 13).Value = (j + 1).ToString();

                        workbookMua.Cell(2 + printIndex2, 14).Value = hddv.ten;
                        workbookMua.Cell(2 + printIndex2, 15).Value = hddv.dvtinh;
                        workbookMua.Cell(2 + printIndex2, 16).Value = hddv.sluong.ToString();
                        workbookMua.Cell(2 + printIndex2, 17).Value = hddv.dgia.ToString();
                        workbookMua.Cell(2 + printIndex2, 18).Value = hddv.tlckhau == null ? "-" : hddv.tlckhau.ToString();
                        workbookMua.Cell(2 + printIndex2, 19).Value = hddv.stckhau == null ? "-" : hddv.stckhau.ToString();
                        workbookMua.Cell(2 + printIndex2, 20).Value = hddv.thtien.ToString();
                        workbookMua.Cell(2 + printIndex2, 21).Value = hddv.tsuat.ToString();
                        workbookMua.Cell(2 + printIndex2, 22).Value = hddv.tsuat == null ? "0" : (hddv.tsuat * hddv.thtien).ToString();
                        workbookMua.Cell(2 + printIndex2, 23).Value = "";

                        // tang index
                        printIndex2++;
                    }
                }

                var filePath = Path.Combine(_saveDirPath, textStartDate.Replace("/", "-") + " đến " + textEndDate.Replace("/", "-") + "_" + Guid.NewGuid().ToString() + ".xlsx");
                using (var stream = new FileStream(filePath, FileMode.OpenOrCreate))
                {
                    workbook.SaveAs(stream);
                }
                richTextBox1.Text += "\nIn ra file thành công!";
            }
            catch(Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
            finally
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }

        void progressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage; //Progress-Value
            object userState = e.UserState; //can be used to pass values to the progress-changed-event
        }
    }
}
