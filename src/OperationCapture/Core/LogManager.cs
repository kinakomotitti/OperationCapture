namespace OperationCapture.Core
{
    using log4net;

    public class LogManager
    {
        public static ILog Logger = log4net.LogManager.GetLogger("rootLogger");
    }
}
