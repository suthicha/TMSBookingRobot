using System;

namespace TMSBookingRobot.Models
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttributes : Attribute
    {
        public string Title { get; set; }
        public int Width { get; set; }
        public int Index { get; set; }
        public bool Exclude { get; set; }
    }
}