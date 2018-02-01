using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TMSBookingRobot.Controllers;
using TMSBookingRobot.Models;
using TMSBookingRobot.Models.Booking;

namespace TMSBookingRobot.TestApp
{
    public partial class Form1 : Form
    {
        private string _sqlConnectionString;
        private BookingController _bookingController;

        public Form1()
        {
            InitializeComponent();
            _sqlConnectionString = ConfigurationManager.AppSettings["Connection"];
            _bookingController = new BookingController();
            _bookingController.SetConnection(_sqlConnectionString);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var bookingQueueObjs = _bookingController.GetBookingQueueFromSysF();

            MessageBox.Show(bookingQueueObjs.Count().ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var bookingNo = "BAE1802-0013";
            var bookingObj = _bookingController.GetBooking(bookingNo);

            MessageBox.Show("OK");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var bookingNo = "BAE1802-0013";
            var status = _bookingController.FlagBookingQueueExportCompleted(bookingNo);

            MessageBox.Show("OK");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var bookingNo = "BAE1802-0013";
            var bookingObj = _bookingController.GetBooking(bookingNo);

            _bookingController.ExportToExcel(bookingObj, bookingNo + ".xlsx");

            //var attrExcelHeaderForBookingControl = new List<ExcelColumnAttributes>();

            //foreach (var prop in bookingObj.GetType().GetProperties())
            //{
            //    if (prop.GetCustomAttributes(typeof(ExcelColumnAttributes), true).Length > 0)
            //    {
            //        var attrs = prop.GetCustomAttributes(true);
            //        var obj = (ExcelColumnAttributes)attrs[0];
            //        obj.Tag = prop.GetValue(bookingObj, null);

            //        attrExcelHeaderForBookingControl.Add(obj);
            //    }
            //}

            //var srtAttrExcelHeaderForBookingControl = attrExcelHeaderForBookingControl.OrderBy(d => d.Index).ToList();
            //var headerTitles = srtAttrExcelHeaderForBookingControl.
            //    Where(q => q.Exclude == false)
            //    .Select(s => s.Title).ToArray();

            //var withCols = srtAttrExcelHeaderForBookingControl.
            //    Where(q => q.Exclude == false)
            //    .Select(s => s.Width).ToArray();

            //var cellValues = srtAttrExcelHeaderForBookingControl.
            //    Where(q => q.Exclude == false)
            //    .Select(s => s.Tag).ToArray();

            //var excelController = new ExcelController();

            //excelController.CreateWorkSheet("CTL");
            //excelController.CreateHeader(headerTitles);
            //excelController.AddRow(cellValues);
            //excelController.AdjustWorksheetContent(withCols);

            //var bookingItems = bookingObj.Items;
            //var bookingItemObj = bookingItems[0];

            //foreach (var prop in bookingItemObj.GetType().GetProperties())
            //{
            //    if (prop.GetCustomAttributes(typeof(ExcelColumnAttributes), true).Length > 0)
            //    {
            //        var attrs = prop.GetCustomAttributes(true);
            //        var obj = (ExcelColumnAttributes)attrs[0];
            //        obj.Tag = prop.GetValue(bookingItemObj, null);

            //        attrExcelHeaderForBookingControl.Add(obj);
            //    }
            //}

            //excelController.CreateWorkSheet("DTL");

            //excelController.Save(bookingNo + ".xlsx");

            MessageBox.Show("Done");
        }
    }
}