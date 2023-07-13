using System;

namespace InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Exceptions
{
    [Serializable]
    public class SpreadsheetGenerationExceptions
    {
        public class OpenFileException : ApplicationException
        {
            public OpenFileException(string message) : base(message) { }
        }

        public class UnknownTypeException : ApplicationException
        {
            public UnknownTypeException(string message) : base(message) { }
        }

        public class CellNotFoundException : ApplicationException
        {
            public CellNotFoundException(string message) : base(message) { }
        }
    }
}