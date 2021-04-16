using System;
using System.Runtime.Serialization;

namespace CableDataParsing.NameBuilders
{
    [Serializable]
    internal class NotImplementException : Exception
    {
        public NotImplementException()
        {
        }

        public NotImplementException(string message) : base(message)
        {
        }

        public NotImplementException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected NotImplementException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}