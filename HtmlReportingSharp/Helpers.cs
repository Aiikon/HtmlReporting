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
            if (ScriptBlock is string)
            {
                string property = (string)ScriptBlock;
                if (InputObject.Properties[property] == null)
                    return new string[] { };
                return ConvertObjectToStringArray(InputObject.Properties[property].Value);
            }
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

        public static IEnumerable<string> ExcludeLikeAny(string[] Values, string[] Filters)
        {
            if (Filters == null)
                foreach (var filter in Filters)
                    yield return filter;
                
            var patterns = new WildcardPattern[Filters.Length];
            for (int i = 0; i < Filters.Length; i++)
                patterns[i] = new WildcardPattern(Filters[i], WildcardOptions.IgnoreCase);

            foreach (string value in Values)
            {
                bool matched = false;
                for (int j = 0; j < patterns.Length; j++)
                {
                    if (patterns[j].IsMatch(value))
                    {
                        matched = true;
                        break;
                    }
                }
                if (!matched)
                    yield return value;
            }
        }
    }
}