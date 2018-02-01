using System;

namespace TMSBookingRobot.Models
{
    public class BaseClass
    {
        public string Key { get { return _key; } }

        private string _key;

        public BaseClass()
        {
            _key = Guid.NewGuid().ToString();
        }
    }
}