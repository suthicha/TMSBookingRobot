using System.IO;

namespace TMSBookingRobot.Controllers
{
    public class BookingAutoBot
    {
        private string _sqlConnection;
        private string _destFolder;

        public BookingAutoBot(string sqlConnectionString, string destFolder)
        {
            _sqlConnection = sqlConnectionString;
            _destFolder = destFolder;
        }

        public void Live()
        {
            var controller = new BookingController();
            controller.SetConnection(_sqlConnection);

            var queue = controller.GetBookingQueueFromSysF();

            if (queue == null || queue.Count == 0) return;

            for (int i = 0; i < queue.Count; i++)
            {
                var bookingQueueItem = queue[i];
                Logger.EventLog("[BookingController.cs][RunAutoBot] : Create job " + bookingQueueItem.BookingNo);

                var bookingJob = controller.GetBooking(bookingQueueItem.BookingNo);

                if (bookingJob == null) continue;
                Logger.EventLog("[BookingController.cs][RunAutoBot] : Find job " + bookingQueueItem.BookingNo);

                var destFileName = Path.Combine(_destFolder, bookingJob + ".xlsx");

                if (controller.ExportToExcel(bookingJob, destFileName))
                {
                    controller.FlagBookingQueueExportCompleted(bookingJob.BookingNo);
                    Logger.EventLog("[BookingController.cs][RunAutoBot] : Write excel " + destFileName);
                }
            }
        }
    }
}