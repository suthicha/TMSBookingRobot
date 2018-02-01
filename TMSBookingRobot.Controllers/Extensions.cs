using System.Data;

namespace TMSBookingRobot.Controllers
{
    public static class Extension
    {
        public static T GetValue<T>(this DataRow row, string columnName, T defaultValue = default(T))
        {
            //object o = row[columnName];
            //if (o is T) return (T)o;
            //return defaultValue;
            return (T)row[columnName];
        }
    }
}