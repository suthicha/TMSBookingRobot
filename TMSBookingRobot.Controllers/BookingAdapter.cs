using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using TMSBookingRobot.Models;
using TMSBookingRobot.Models.Booking;

namespace TMSBookingRobot.Controllers
{
    internal class BookingAdapter
    {
        internal List<BookingQueue> ConvertDataSetToBookingQueueItems(DataSet data)
        {
            if (data == null || data.Tables[0].Rows.Count == 0)
                return null;

            var bookingQueueItems = new List<BookingQueue>();

            for (int i = 0; i < data.Tables[0].Rows.Count; i++)
            {
                DataRow dr = data.Tables[0].Rows[i];

                try
                {
                    bookingQueueItems.Add(new BookingQueue()
                    {
                        TrxNo = dr.GetValue<int>("TrxNo"),
                        BookingNo = dr.GetValue<string>("BookingNo"),
                        ShipperCode = dr.GetValue<string>("ShipperCode")
                    });
                }
                catch (Exception ex)
                {
                    Logger.ErrorLog("[BookingAdapter.cs][ConvertDataSetToBookingQueueItems] " + ex.Message);
                }
            }

            return bookingQueueItems;
        }

        internal List<BookingItem> ConvertDataTableToBookingItems(DataTable table)
        {
            var bookingItems = new List<BookingItem>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                DataRow dr = table.Rows[i];
                try
                {
                    bookingItems.Add(new BookingItem()
                    {
                        BookingNo = dr.GetValue<string>("BookingNo"),
                        CommercialInvoiceNo = dr.GetValue<string>("CommercialInvoiceNo"),
                        ProductType = dr.GetValue<string>("ProductType"),
                        PartNo = dr.GetValue<string>("PartNo"),
                        ProductDescription = dr.GetValue<string>("ProductDescription"),
                        Quantity = dr.GetValue<decimal>("QTY"),
                        Unit = dr.GetValue<string>("Unit"),
                        Weight = dr.GetValue<decimal>("Weight"),
                        WeightUnit = dr.GetValue<string>("WeightUnit"),
                        Width = dr.GetValue<decimal>("Width"),
                        Length = dr.GetValue<decimal>("Length"),
                        Height = dr.GetValue<decimal>("Height"),
                        LoadFromLocation = dr.GetValue<string>("LoadFromLocation"),
                        LoadDate = dr.GetValue<DateTime>("LoadDate"),
                        DropToLocation = dr.GetValue<string>("DropToLocation"),
                        DropDate = dr.GetValue<DateTime>("DropDate"),
                        Volume = dr.GetValue<decimal>("Volume")
                    });
                }
                catch (Exception ex)
                {
                    Logger.ErrorLog("[BookingAdapter.cs][ConvertDataTableToBookingItems] " + ex.Message);
                }
            }

            return bookingItems;
        }

        internal BookingControl ConvertDataSetToBookingControl(DataSet data)
        {
            BookingControl bookingControlObj = null;

            if (data == null || data.Tables.Count == 1) return bookingControlObj;

            try
            {
                var table = data.Tables[0];
                DataRow dr = table.Rows[0];

                bookingControlObj = new BookingControl()
                {
                    BookingNo = dr.GetValue<string>("BookingNo"),
                    CustomerCode = dr.GetValue<string>("CustomerCode"),
                    CustomerBranch = dr.GetValue<string>("CustomerBranch"),
                    ShipperCode = dr.GetValue<string>("ShipperCode"),
                    ShipperBranch = dr.GetValue<string>("ShipperBranch"),
                    EndCustomerCode = dr.GetValue<string>("EndCustomerCode"),
                    EndCustomerBranch = dr.GetValue<string>("EndCustomerBranch"),
                    Currency = dr.GetValue<string>("Currency"),
                    BookingType = dr.GetValue<string>("BookingType"),
                    ServiceType = dr.GetValue<string>("ServiceType"),
                    BookingOffice = dr.GetValue<string>("BookingOffice"),
                    RouteFrom = dr.GetValue<string>("RouteFrom"),
                    RouteTo = dr.GetValue<string>("RouteTo"),
                    Email = dr.GetValue<string>("Email")
                };

                var bookingItemObjs = ConvertDataTableToBookingItems(data.Tables[1]);
                bookingControlObj.AddItems(bookingItemObjs.ToArray());
            }
            catch (Exception ex) { Logger.ErrorLog("[BookingAdapter.cs][ConvertDataSetToBookingControl] " + ex.Message); }

            return bookingControlObj;
        }

        internal List<ExcelColumnAttributes> GetExcelColumnAttributes(Object classObj)
        {
            var attrExcelColumns = new List<ExcelColumnAttributes>();

            foreach (var prop in classObj.GetType().GetProperties())
            {
                if (prop.GetCustomAttributes(typeof(ExcelColumnAttributes), true).Length > 0)
                {
                    var attrs = prop.GetCustomAttributes(true);
                    var obj = (ExcelColumnAttributes)attrs[0];

                    if (prop.PropertyType == typeof(DateTime))
                    {
                        var dateVal = prop.GetValue(classObj, null);
                        obj.Tag = ((DateTime)dateVal).ToString("dd-MMM-yyyy HH:mm", new CultureInfo("en-US"));
                        obj.CellType = "DateTime";
                    }
                    else
                    {
                        obj.Tag = prop.GetValue(classObj, null);
                        obj.CellType = "Text";
                    }

                    attrExcelColumns.Add(obj);
                }
            }

            return attrExcelColumns;
        }
    }
}