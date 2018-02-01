using System.Collections.Generic;
using System.Linq;

namespace TMSBookingRobot.Models.Booking
{
    public class BookingControl : BaseClass
    {
        [ExcelColumnAttributes(
            Exclude = true,
            Index = 0,
            Title = "Booking No.",
            Width = 20
         )]
        public string BookingNo { get; set; }

        [ExcelColumnAttributes(
            Index = 1,
            Title = "CustomerCode",
            Width = 20
         )]
        public string CustomerCode { get; set; }

        [ExcelColumnAttributes(
            Index = 2,
            Title = "CustomerBranch",
            Width = 15
         )]
        public string CustomerBranch { get; set; }

        [ExcelColumnAttributes(
            Index = 3,
            Title = "ShipperCode",
            Width = 20
         )]
        public string ShipperCode { get; set; }

        [ExcelColumnAttributes(
            Index = 4,
            Title = "ShipperBranch",
            Width = 15
         )]
        public string ShipperBranch { get; set; }

        [ExcelColumnAttributes(
            Index = 5,
            Title = "EndCustomerCode",
            Width = 20
         )]
        public string EndCustomerCode { get; set; }

        [ExcelColumnAttributes(
            Index = 6,
            Title = "EndCustomerBranch",
            Width = 20
         )]
        public string EndCustomerBranch { get; set; }

        [ExcelColumnAttributes(
            Index = 7,
            Title = "Currency",
            Width = 10
         )]
        public string Currency { get; set; }

        [ExcelColumnAttributes(
            Index = 8,
            Title = "BookingType",
            Width = 10
         )]
        public string BookingType { get; set; }

        [ExcelColumnAttributes(
            Index = 9,
            Title = "ServiceType",
            Width = 20
         )]
        public string ServiceType { get; set; }

        [ExcelColumnAttributes(
            Index = 10,
            Title = "BookingOffice",
            Width = 20
         )]
        public string BookingOffice { get; set; }

        [ExcelColumnAttributes(
            Index = 11,
            Title = "RouteFrom",
            Width = 20
         )]
        public string RouteFrom { get; set; }

        [ExcelColumnAttributes(
            Index = 12,
            Title = "RouteTo",
            Width = 20
         )]
        public string RouteTo { get; set; }

        [ExcelColumnAttributes(
            Index = 13,
            Title = "Email",
            Width = 20
         )]
        public string Email { get; set; }

        public List<BookingItem> Items { get { return _bookingItems; } }

        private List<BookingItem> _bookingItems;

        public BookingControl()
        {
            _bookingItems = new List<BookingItem>();
        }

        public int ItemCount()
        {
            return Items.Count;
        }

        public void AddItems(params BookingItem[] items)
        {
            _bookingItems.AddRange(items);
        }

        public void AddItem(BookingItem item)
        {
            _bookingItems.Add(item);
        }

        public void DeleteItem(BookingItem item)
        {
            var itemToRemove = _bookingItems.Single(q => q.Key == item.Key);
            _bookingItems.Remove(itemToRemove);
        }
    }
}