using System;

namespace ConfigLoader
{
    public class ConfigLoaderException : Exception
    {
        public String OriginalMessage { get; private set; }

        public ConfigLoaderException(String message, String originalMessage)
            : base(message)
        {
            OriginalMessage = originalMessage;
        }
    }
}
