namespace TMSBookingRobot.Models.Booking
{
    public class BookingQueue : BaseClass
    {
        public int TrxNo { get; set; }
        public string BookingNo { get; set; }
        public string ShipperCode { get; set; }
    }
}