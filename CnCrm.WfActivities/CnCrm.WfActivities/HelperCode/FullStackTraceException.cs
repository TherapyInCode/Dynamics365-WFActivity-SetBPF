namespace CnCrm.WfActivities.HelperCode {
  using System;
  using System.ServiceModel;

  /// <summary>
  /// Provides a method for handling any exception types that "eat" the stacktrace in it's implementation of ToString()
  /// </summary>
  [Serializable]
  public class FullStackTraceException : Exception {
    private readonly Exception exception;

    private FullStackTraceException(Exception ex) {
      this.exception = ex;
    }

    public static Exception Create(Exception exception) {
      return CreateInternal((dynamic)exception);
    }

    private static Exception CreateInternal(Exception exception) {
      return exception;
    }

    public static FullStackTraceException Create<TDetail>(FaultException<TDetail> exception) {
      return new FullStackTraceException(exception);
    }

    private static FullStackTraceException CreateInternal<TDetail>(FaultException<TDetail> exception) {
      return new FullStackTraceException(exception);
    }

    public override String ToString() {
      var s = this.exception.ToString();
      if (this.exception.InnerException != null) {
        s = String.Format("{0} ---> {1}{2}   --- End of inner exception stack trace ---{2}", s, this.exception.InnerException, Environment.NewLine);
      }

      var stackTrace = this.exception.StackTrace;
      if (stackTrace != null) {
        s += Environment.NewLine + stackTrace;
      }

      return s;
    }
  }
}