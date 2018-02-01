using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using TMSBookingRobot.Models.Booking;

namespace TMSBookingRobot.Controllers
{
    public class BookingController
    {
        private string _sqlConnectionString;

        public BookingController()
        {
        }

        public void SetConnection(string sqlConnectionString)
        {
            this._sqlConnectionString = sqlConnectionString;
        }

        public List<BookingQueue> GetBookingQueueFromSysF()
        {
            DataSet dsBookingQueue = DbAdapter.QueryToDS(
                @"SELECT TrxNo, BookingNo, ShipperCode FROM tms_booking 
                WHERE DocStatus = '1' 
                AND ShipperCode LIKE 'NMB%' 
                GROUP BY TrxNo, BookingNo, ShipperCode", _sqlConnectionString);

            var bookingAdapter = new BookingAdapter();
            return bookingAdapter.ConvertDataSetToBookingQueueItems(dsBookingQueue);
        }

        public BookingControl GetBooking(string bookingNo)
        {
            DataSet dsBooking = DbAdapter.QueryToDS(
                                "sp_cti_BookingAirExportToTMS",
                                _sqlConnectionString,
                                CommandType.StoredProcedure,
                                new SqlParameter[] {
                                    new SqlParameter("@BookingNo", bookingNo)
                                });

            var bookingAdapter = new BookingAdapter();
            return bookingAdapter.ConvertDataSetToBookingControl(dsBooking);
        }

        public int FlagBookingQueueExportCompleted(string bookingNo)
        {
            return DbAdapter.Executed(@"UPDATE tms_booking SET DocStatus='2', UpdateDateTime=GETDATE() 
                WHERE BookingNo=@BookingNo",
                _sqlConnectionString,
                CommandType.Text,
                new SqlParameter[] {
                    new SqlParameter("@BookingNo", bookingNo)
                });
        }

        public bool ExportToExcel(BookingControl booking, string destinationFileName)
        {
            if (File.Exists(destinationFileName)) File.Exists(destinationFileName);

            var bookingAdapter = new BookingAdapter();
            var attrExcelHeaderForBookingControl = bookingAdapter.GetExcelColumnAttributes(booking);

            var srtAttrExcel = attrExcelHeaderForBookingControl.OrderBy(d => d.Index).ToList();
            var headerTitles = srtAttrExcel.
                Where(q => q.Exclude == false)
                .Select(s => s.Title).ToArray();

            var withCols = srtAttrExcel.
                Where(q => q.Exclude == false)
                .Select(s => s.Width).ToArray();

            var cellValues = srtAttrExcel.
                Where(q => q.Exclude == false)
                .Select(s => new { Tag = s.Tag, CellType = s.CellType }).ToArray();

            var excelController = new ExcelController();

            excelController.CreateWorkSheet("CTL");
            excelController.CreateHeader(headerTitles);
            excelController.AddRow(cellValues);
            excelController.AdjustWorksheetContent(withCols);

            if (booking.ItemCount() > 0)
            {
                var attrExcelHeaderForBookingItem = bookingAdapter.GetExcelColumnAttributes(booking.Items[0]);

                srtAttrExcel = attrExcelHeaderForBookingItem.OrderBy(d => d.Index).ToList();
                headerTitles = srtAttrExcel.
                    Where(q => q.Exclude == false)
                    .Select(s => s.Title).ToArray();

                withCols = srtAttrExcel.
                    Where(q => q.Exclude == false)
                    .Select(s => s.Width).ToArray();

                excelController.CreateWorkSheet("DTL");
                excelController.CreateHeader(headerTitles);

                for (int i = 0; i < booking.ItemCount(); i++)
                {
                    var bookingItemObj = booking.Items[i];
                    attrExcelHeaderForBookingItem = bookingAdapter.GetExcelColumnAttributes(bookingItemObj);
                    srtAttrExcel = attrExcelHeaderForBookingItem.OrderBy(d => d.Index).ToList();

                    cellValues = srtAttrExcel.
                        Where(q => q.Exclude == false)
                        .Select(s => new { Tag = s.Tag, CellType = s.CellType }).ToArray();

                    excelController.AddRow(cellValues);
                }
                excelController.AdjustWorksheetContent(withCols);
            }

            excelController.Save(destinationFileName);

            return File.Exists(destinationFileName);
        }
    }
}