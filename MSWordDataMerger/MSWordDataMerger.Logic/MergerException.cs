using System;

namespace MSWordDataMerger.Logic
{
    public class MergerException : Exception
    {
        public String OriginalMessage { get; private set; }
        public MergerException(String message, String originalMessage) 
            : base(message)
        {
            OriginalMessage = originalMessage;
        }
    }
}
