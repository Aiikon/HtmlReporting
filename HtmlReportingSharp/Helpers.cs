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

        public static string[] InvokeScriptblockReturnStringArray(object ScriptBlock, PSObject InputObject)
        {
            if (ScriptBlock == null)
                return new string[] { };
            ScriptBlock scriptBlock = (ScriptBlock)ScriptBlock;
            string[] result = scriptBlock.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", InputObject) }, null)
                .Where(r => r != null)
                .Select(r => r.ToString())
                .ToArray();
            return result;
        }

        public static string[] ConvertObjectToStringArray(object Object)
        {
            if (Object == null)
                return new string[] { };
            if (Object is Array)
                return ((object[])Object).Where(o => o != null).Select(o => o.ToString()).ToArray();
            return new string[] { Object.ToString() };
        }
    }
}