using log4net;

namespace OperationCapture.Core
{
    public class LogManager
    {
        public static ILog Logger = log4net.LogManager.GetLogger("rootLogger");
    }
}
