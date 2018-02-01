using System;

namespace TMSBookingRobot.Models.Booking
{
    public class BookingItem
    {
        [ExcelColumnAttributes(
            Exclude = true,
            Index = 0,
            Title = "Booking No.",
            Width = 15
         )]
        public string BookingNo { get; set; }

        [ExcelColumnAttributes(
            Index = 1,
            Title = "CommercialInvoiceNo",
            Width = 20
         )]
        public string CommercialInvoiceNo { get; set; }

        [ExcelColumnAttributes(
            Index = 2,
            Title = "ProductType",
            Width = 10
         )]
        public string ProductType { get; set; }

        [ExcelColumnAttributes(
            Index = 3,
            Title = "Part No.",
            Width = 15
        )]
        public string PartNo { get; set; }

        [ExcelColumnAttributes(
            Index = 4,
            Title = "ProductDescription",
            Width = 30
        )]
        public string ProductDescription { get; set; }

        [ExcelColumnAttributes(
            Index = 5,
            Title = "QTY",
            Width = 10
         )]
        public decimal Quantity { get; set; }

        [ExcelColumnAttributes(
            Index = 6,
            Title = "Unit",
            Width = 5
         )]
        public string Unit { get; set; }

        [ExcelColumnAttributes(
            Index = 7,
            Title = "Weight",
            Width = 10
         )]
        public decimal Weight { get; set; }

        [ExcelColumnAttributes(
            Index = 8,
            Title = "WeightUnit",
            Width = 10
         )]
        public string WeightUnit { get; set; }

        [ExcelColumnAttributes(
            Index = 9,
            Title = "Width",
            Width = 10
         )]
        public decimal Width { get; set; }

        [ExcelColumnAttributes(
            Index = 10,
            Title = "Length",
            Width = 10
         )]
        public decimal Length { get; set; }

        [ExcelColumnAttributes(
            Index = 11,
            Title = "Height",
            Width = 10
         )]
        public decimal Height { get; set; }

        [ExcelColumnAttributes(
            Index = 12,
            Title = "LoadFromLocation",
            Width = 30
         )]
        public string LoadFromLocation { get; set; }

        [ExcelColumnAttributes(
            Index = 13,
            Title = "LoadDate",
            Width = 20
         )]
        public DateTime LoadDate { get; set; }

        [ExcelColumnAttributes(
            Index = 14,
            Title = "DropToLocation",
            Width = 30
         )]
        public string DropToLocation { get; set; }

        [ExcelColumnAttributes(
            Index = 15,
            Title = "DropDate",
            Width = 20
         )]
        public DateTime DropDate { get; set; }

        [ExcelColumnAttributes(
            Index = 16,
            Title = "Volume",
            Width = 10
         )]
        public decimal Volume { get; set; }
    }
}