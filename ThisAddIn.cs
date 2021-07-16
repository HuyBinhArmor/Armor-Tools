using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Windows.Forms;




namespace ArmorAcc
{
    public partial class ThisAddIn //Class Hàm hệ thống 
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Workbook GetActiveWorkbook()
        {
            return (Excel.Workbook)Application.ActiveWorkbook;
        }

        public Excel.Worksheet GetActiveSheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        public Excel.Range GetSelection()
        {
            return (Excel.Range)Application.Selection;
        }

        public Excel.Range GetLastCellFilled()
        {
            return Globals.ThisAddIn.GetActiveSheet().Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        }





        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class EventOpenClose : IExcelAddIn //Class Even Open & Clode
    {
        public void AutoOpen()
        {
            // Versions before v1.1.0 required only a call to Register() in the AutoOpen().
            // The name was changed (and made obsolete) to highlight the pair of function calls now required.
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    } //Class Even Open & Clode

    public class Excel_UDF_DOCSOTIEN //Tạo hàm đọc số tiền bằng chữ 
    {
        private static string[] m09Text = new string[10] { " không", " một", " hai", " ba", " bốn", " năm", " sáu", " bảy", " tám", " chín" };
        private static string[] mHauto = new string[6] { "", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ" };
        // Hàm đọc bộ 3 số
        private static string BobasoText(int _BoBaSo)
        {
            int _SoHangTram, _SoHangChuc, _SoHangDvi;
            string _mString = "";
            _SoHangTram = (int)(_BoBaSo / 100);
            _SoHangChuc = (int)((_BoBaSo % 100) / 10);
            _SoHangDvi = _BoBaSo % 10;

            if ((_SoHangTram == 0) && (_SoHangChuc == 0) && (_SoHangDvi == 0)) return "";
            // Xét số hàng trăm
            _mString += m09Text[_SoHangTram] + " trăm";
            // Xét số hàng chục
            if ((_SoHangChuc == 0) && (_SoHangDvi == 0)) return _mString;
            if ((_SoHangChuc == 0) && (_SoHangDvi > 0)) _mString += " linh";
            if (_SoHangChuc == 1) _mString += " mười";
            if (_SoHangChuc > 1) _mString += m09Text[_SoHangChuc] + " mươi";
            // Xét số hàng đơn vị
            switch (_SoHangDvi)
            {
                case 1:
                    if (_SoHangChuc > 1)
                    {
                        _mString += " mốt";
                    }
                    else
                    {
                        _mString += " một";
                    }
                    break;
                case 5:
                    if (_SoHangChuc == 0)
                    {
                        _mString += " năm";
                    }
                    else
                    {
                        _mString += " lăm";
                    }
                    break;
                default:
                    if (_SoHangDvi != 0)
                    {
                        _mString += m09Text[_SoHangDvi];
                    }
                    break;
            }

            return _mString;
        }
        // Hàm đọc số bằng chữ
        [ExcelFunction(Description = "Đọc số tiền bằng chữ")]
        public static string BangChu(
            [ExcelArgument(Name = " (*) Number", Description = " Số cần đọc thành chữ")] long _SoTien,
            [ExcelArgument(Name = " (-) ĐVT", Description = " Đvt, Nhớ để trong dấu ngoặc nháy '' '' haha")] string _DVT)
        {
            int _SoLop, i;
            string _Dau = "";
            string _mString = "", _Boba = "";
            int[] Lop = new int[6];
            if (_SoTien == 0) return "Không";
            if (_SoTien < 0)
            {
                _SoTien = -_SoTien;
                _Dau = "Âm ";
            }

            //Kiểm tra số quá lớn
            if (_SoTien > 9000000000000000)
            {
                return "";
            }
            Lop[5] = (int)(_SoTien / Math.Pow(10, 15));
            _SoTien = (long)(_SoTien % Math.Pow(10, 15));

            Lop[4] = (int)(_SoTien / Math.Pow(10, 12));
            _SoTien = (long)(_SoTien % Math.Pow(10, 12));

            Lop[3] = (int)(_SoTien / Math.Pow(10, 9));
            _SoTien = (long)(_SoTien % Math.Pow(10, 9));

            Lop[2] = (int)(_SoTien / Math.Pow(10, 6));
            _SoTien = (long)(_SoTien % Math.Pow(10, 6));

            Lop[1] = (int)(_SoTien / Math.Pow(10, 3));
            _SoTien = (long)(_SoTien % Math.Pow(10, 3));

            Lop[0] = (int)(_SoTien / Math.Pow(10, 0));




            if (Lop[5] > 0)
            {
                _SoLop = 5;
            }
            else if (Lop[4] > 0)
            {
                _SoLop = 4;
            }
            else if (Lop[3] > 0)
            {
                _SoLop = 3;
            }
            else if (Lop[2] > 0)
            {
                _SoLop = 2;
            }
            else if (Lop[1] > 0)
            {
                _SoLop = 1;
            }
            else
            {
                _SoLop = 0;
            }
            for (i = _SoLop; i >= 0; i--)
            {
                _Boba = BobasoText(Lop[i]);
                _mString += _Boba;

                if (Lop[i] != 0) _mString += mHauto[i];
                if ((i > 0) && (!string.IsNullOrEmpty(_Boba))) _mString += ",";
            }


            if (_mString.Substring(0, 16) == " không trăm linh")
            {
                _mString = _mString.Substring(16, _mString.Length - 16);
            }
            else if (_mString.Substring(0, 11) == " không trăm")
            {
                _mString = _mString.Substring(11, _mString.Length - 11);
            }

            _mString = _Dau + _mString.Trim() + " " + _DVT; // Thêm dấu âm và đơn vị tính
            return _mString.Substring(0, 1).ToUpper() + _mString.Substring(1); // Viết hoa chữ đầu

        }
    }

    public class Excel_UDF_TNCN // Tạo hàm tính thuế TNCN 
    {
        [ExcelFunction(Description = "Tính thuế TNCN khấu trừ theo biểu lũy tiến")]
        public static double PIT
            (
            [ExcelArgument(Name = "(*) TN Tính thuế", Description = " TNTT = (Tổng TN)-(TN miễn thuế)-(BHXH)-(GTGC)-(GT khác)")] double _TNTT)
        {
            double _TNCN;
            if (_TNTT <= 0)
            {
                _TNCN = 0;
            }
            else if (_TNTT <= 5000000)
            {
                _TNCN = _TNTT * 5 / 100;
            }
            else if (_TNTT <= 10000000)
            {
                _TNCN = _TNTT * 10 / 100 - 250000;
            }
            else if (_TNTT <= 18000000)
            {
                _TNCN = _TNTT * 15 / 100 - 750000;
            }
            else if (_TNTT <= 32000000)
            {
                _TNCN = _TNTT * 20 / 100 - 1650000;
            }
            else if (_TNTT <= 52000000)
            {
                _TNCN = _TNTT * 25 / 100 - 3250000;
            }
            else if (_TNTT <= 80000000)
            {
                _TNCN = _TNTT * 30 / 100 - 5850000;
            }
            else
            {
                _TNCN = _TNTT * 35 / 100 - 9850000;
            }
            return Math.Round(_TNCN, 0);
        }
    }

    public class Excel_UDF_LOOKUPPIC 
    {
        // Chuyyển định dạng từ ExcelReference thành Excel.Range
        private static Excel.Range ReferenceToRange(ExcelReference xlRef)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            string strAddress = XlCall.Excel(XlCall.xlfReftext, xlRef, true).ToString();
            return app.Range[strAddress];
        }

        [ExcelFunction(Description = "Lấy hình ảnh từ Folder Path", IsMacroType = true)]
        public static string SearchPic(
            [ExcelArgument(Name = "(*) FilePath", Description = "Đường dẫn của ảnh")] string Filepath,
            [ExcelArgument(Name = "(*) Shape Range", Description = "Ô cần đặt Ảnh kết quả", AllowReference = true)] object Cells)
        {
            if (Filepath == "") return "Quên nhập Url";

            ExcelReference rng;
            // Chuyển định dạng
            try
            {
                rng = (ExcelReference)Cells;
            }
            catch (Exception NullRange)
            {
                return " Quên nhập Cell trả kết quả";
            }
            Excel.Range mRange = ReferenceToRange(rng);

            // Khai báo biến
            Excel.Worksheet ActiveSheet = mRange.Worksheet;
            Excel.Shape mPicture;
            String mPictureName = "Search Pic _ " + ActiveSheet.Name.ToString() + mRange.Address.ToString();

            foreach (Excel.Shape DelShape in ActiveSheet.Shapes)
            {
                if (DelShape.Name == mPictureName)
                {
                    DelShape.Delete();
                }
            }

            // Add Picture
            try
            {
                mPicture = ActiveSheet.Shapes.AddPicture(Filepath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, mRange.Left + 3, mRange.Top + 3, -1, -1);
                mPicture.Name = mPictureName;
            }
            catch (Exception NA)
            {
                return " Không tìm thấy ảnh";
            }
           
            // Resize Auto fit
            if ((mRange.Width / mRange.Height) > (mPicture.Width / mPicture.Height))
            {
                mPicture.Height = Convert.ToSingle(mRange.Height) - 3;
                mPicture.Width -= 3;
            }
            else
            {
                mPicture.Width = Convert.ToSingle(mRange.Width) - 3;
                mPicture.Height -= 3;
            }
            return mPictureName;
        }


    } // Tạo hàm tìm kiếm từ File Path

    public class Excel_UDF_GetQRcodeUrl 
    {
        [ExcelFunction(Description = " Tạo link QR code từ chuỗi String")]
        public static string GetQRcodeUrl(
            [ExcelArgument(Name = "(*) Input String", Description = " Dòng chữ cần tạo QR code")] string InputString,
            [ExcelArgument(Name = "(-) C.Rộng", Description = " Mặc định là 100")] int width,
            [ExcelArgument(Name = "(-) C.Cao", Description = " Mặc định là 100")] int height,
            [ExcelArgument(Name = "(-) Encoding Option", Description = " Mặc định là UTF")] string encoding,
            [ExcelArgument(Name = "(-) Errorlevel", Description = " Chọn {L,M,Q,H} Mặc định là L")] string errorlevel,
            [ExcelArgument(Name = "(-) Canh Lề",Description =" Mặc định là 0")] int margin)
        {
            if (InputString == "") return "Chưa nhập Input String";
            // Set Defaults
            if (width < 100) width = 100;
            if (height < 100) height = 100;
            if (encoding == "") encoding = "UTF-8";
            if (errorlevel == "") errorlevel = "L";
            if (margin <= 0) margin = 0;

            string Link = "https://chart.googleapis.com/chart?cht=qr&chl=QRDATA&chs=WIDTHxHEIGHT&choe=ENCODING&chld=ERRORCORRECTIONLEVEL|MARGIN";

            Link = Link.Replace("QRDATA", InputString);
            Link = Link.Replace("WIDTH", width.ToString());
            Link = Link.Replace("HEIGHT", height.ToString());
            Link = Link.Replace("ENCODING", encoding);
            Link = Link.Replace("ERRORCORRECTIONLEVEL", errorlevel);
            Link = Link.Replace("MARGIN", margin.ToString());
 
            return Link;
        }


    }  // Taọ hàm lấy QR code Url



}
