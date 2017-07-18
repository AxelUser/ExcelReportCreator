using System;

namespace ExcelReportsCreator
{

    [Serializable]
    public class ReportBuilderException : Exception
    {
        public ReportBuilderException() { }
        public ReportBuilderException(string message) : base(message) { }
        public ReportBuilderException(string message, Exception inner) : base(message, inner) { }
        protected ReportBuilderException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
