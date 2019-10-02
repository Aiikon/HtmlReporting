using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;

namespace HtmlReportingSharp
{
    class Helpers
    {
        public static ErrorRecord NewInvalidDataErrorRecord(string Message)
        {
            return new ErrorRecord(new Exception(Message), "", ErrorCategory.InvalidData, null);
        }
    }
}