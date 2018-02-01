using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using TMSBookingRobot.Models;
using TMSBookingRobot.Models.Booking;

namespace TMSBookingRobot.Controllers
{
    public class BookingController
    {
        private string _sqlConnectionString;

        public BookingController()
        {
            Logger.EventLog("init BookingController");
        }

        public void SetConnection(string sqlConnectionString)
        {
            this._sqlConnectionString = sqlConnectionString;
        }

        public List<BookingQueue> GetBookingQueueFromSysF()
        {
            DataSet dsBookingQueue = DbAdapter.QueryToDS(
                @"SELECT TrxNo, BookingNo, ShipperCode FROM tms_booking 
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

        public int DeleteBookingQueue(string bookingNo)
        {
            return DbAdapter.Executed(@"DELETE tms_booking WHERE BookingNo=@BookingNo",
                _sqlConnectionString,
                CommandType.Text,
                new SqlParameter[] {
                    new SqlParameter("@BookingNo", bookingNo)
                });
        }
    }
}