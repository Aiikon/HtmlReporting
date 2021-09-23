using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections;

namespace HtmlReportingSharp
{
    [Cmdlet(VerbsData.ConvertTo, "HtmlTable")]
    public class ConvertToHtmlTable : PSCmdlet
    {
        [Parameter(ValueFromPipeline = true)]
        public PSObject InputObject { get; set; }

        [Parameter()]
        public string[] Property { get; set; }

        [Parameter()]
        public string[] ExcludeProperty { get; set; }

        [Parameter()]
        public string[] HtmlProperty { get; set; }

        [Parameter()]
        public string[] Class { get; set; }

        [Parameter()]
        public string Id { get; set; }

        [Parameter()]
        public SwitchParameter RowsOnly { get; set; }

        [Parameter()]
        public ScriptBlock RowClassScript { get; set; }

        [Parameter()]
        public ScriptBlock RowStyleScript { get; set; }

        [Parameter()]
        public Hashtable CellClassScripts { get; set; }

        [Parameter()]
        public Hashtable CellStyleScripts { get; set; }

        [Parameter()]
        public Hashtable CellColspanScripts { get; set; }

        [Parameter()]
        public Hashtable CellRowspanScripts { get; set; }

        [Parameter()]
        public Hashtable RenameHeader { get; set; }

        [Parameter()]
        public string[] RightAlignProperty { get; set; }

        [Parameter()]
        public string[] NoWrapProperty { get; set; }

        [Parameter()]
        public SwitchParameter Narrow { get; set; }

        [Parameter()]
        public SwitchParameter Plain { get; set; }

        [Parameter()]
        public SwitchParameter AutoDetectHtml { get; set; }

        [Parameter()]
        public SwitchParameter AddDataColumnName { get; set; }

        [Parameter()]
        public string NoContentHtml { get; set; }

        [Parameter()]
        public string[] InsertSolidLine { get; set; }

        [Parameter()]
        public string[] InsertDashedLine { get; set; }

        [Parameter()]
        public string[] InsertDottedLine { get; set; }
        
        private List<PSObject> inputObjectList = new List<PSObject>();

        private int GetCountFromProperty(PSObject inputObject, Hashtable dictionary, string header)
        {
            if (!dictionary.ContainsKey(header))
                return 0;
            object scriptOrString = dictionary[header];
            if (scriptOrString is ScriptBlock)
            {
                var script = (ScriptBlock)scriptOrString;
                var result = script.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", inputObject) }, null);
                if (result.Count == 0)
                    return 0;
                else if (result.Count > 1)
                {
                    Helpers.NewInvalidDataErrorRecord(String.Format("Count ScriptBlock for {0} returned more than one object.", header));
                    return 0;
                }
                var value = result[0];
                try
                {
                    return (int)value.BaseObject;
                }
                catch
                {
                    Helpers.NewInvalidDataErrorRecord(String.Format("Count ScriptBlock for {0} returned a non-integer object: {1}.", header, value));
                    return 0;
                }
            }
            else if (scriptOrString is String)
            {
                string property = scriptOrString.ToString();
                if (inputObject.Properties[header] == null)
                    return 0;
                var psProperty = inputObject.Properties[header];
                if (psProperty.Value == null)
                    return 0;
                try
                {
                    return (int)psProperty.Value;
                }
                catch
                {
                    Helpers.NewInvalidDataErrorRecord(String.Format("Count Property {0} for {1} could not be cast as an integer: {1}.", property, header, psProperty.Value));
                    return 0;
                }
            }
            else
            {
                Helpers.NewInvalidDataErrorRecord(String.Format("Count object for {0} was not a ScriptBlock or String", header));
                return 0;
            }
        }

        protected override void ProcessRecord()
        {
            if (InputObject != null)
                inputObjectList.Add(InputObject);
        }

        protected override void EndProcessing()
        {
            var resultBuilder = new StringBuilder();
            var headerList = new List<string>();

            if (HtmlProperty == null)
                HtmlProperty = new string[] { };

            if (NoWrapProperty == null)
                NoWrapProperty = new string[] { };

            if (RightAlignProperty == null)
                RightAlignProperty = new string[] { };

            bool noWrapAll = NoWrapProperty.Length == 1 && NoWrapProperty[0] == "*";

            if (Property == null && inputObjectList.Count != 0)
                Property = inputObjectList[0].Properties.Select(p => p.Name).ToArray();

            if (ExcludeProperty != null)
                Property = Helpers.ExcludeLikeAny(Property, ExcludeProperty).ToArray();
            
            if (inputObjectList.Count != 0)
                foreach (var p in Property)
                    headerList.Add(p);

            var renameHeaderDict = new Dictionary<string,string>();
            if (RenameHeader != null)
                foreach (object key in RenameHeader.Keys)
                    if (RenameHeader[key] != null)
                        renameHeaderDict[key.ToString()] = RenameHeader[key].ToString();

            var lineClassDict = new Dictionary<string,string>();
            if (InsertSolidLine != null)
                foreach (string p in InsertSolidLine)
                    lineClassDict[p] = "Solid";
            if (InsertDashedLine != null)
                foreach (string p in InsertDashedLine)
                    lineClassDict[p] = "Dashed";
            if (InsertDottedLine != null)
                foreach (string p in InsertDottedLine)
                    lineClassDict[p] = "Dotted";

            if (!RowsOnly.IsPresent && inputObjectList.Count != 0)
            {
                var classList = new List<string>();
                if (!Plain.IsPresent)
                    classList.Add("HtmlReportingTable");
                if (Narrow.IsPresent)
                    classList.Add("Narrow");
                if (Class != null)
                    foreach (var c in Class)
                        classList.Add(c);    
                resultBuilder.AppendFormat("<table class='{0}'", String.Join(" ", classList));
                if (Id != null)
                    resultBuilder.AppendFormat(" id='{0}'", Id);
                resultBuilder.Append(">\r\n<thead>\r\n");
                resultBuilder.Append("<tr class='header'>\r\n");
                foreach (string header in headerList)
                {
                    resultBuilder.Append("<th");
                    if (AddDataColumnName.IsPresent)
                        resultBuilder.AppendFormat(" data-column-name='{0}'", System.Web.HttpUtility.HtmlAttributeEncode(header));
                    if (lineClassDict.ContainsKey(header))
                        resultBuilder.AppendFormat(" class='Insert{0}Line'", lineClassDict[header]);
                    if (renameHeaderDict.ContainsKey(header))
                        resultBuilder.Append(String.Format(">{0}</th>\r\n", renameHeaderDict[header]));
                    else
                        resultBuilder.Append(String.Format(">{0}</th>\r\n", header));
                }
                resultBuilder.Append("</tr>\r\n");
                resultBuilder.Append("</thead>\r\n");
                resultBuilder.Append("<tbody>\r\n");
            }
            
            if (inputObjectList.Count == 0 && NoContentHtml != null)
                resultBuilder.Append(NoContentHtml);

            var rowspanCountHash = new Dictionary<string, int>();
            var colspanRowHash = new Dictionary<string, int>();
            var colspanCountHash = new Dictionary<string, int>();

            foreach (var inputObject in inputObjectList)
            {
                var rowClassList = new List<string>();
                var rowStyleList = new List<string>();

                rowClassList.AddRange(Helpers.InvokeScriptblockReturnStringArray(RowClassScript, inputObject));
                rowStyleList.AddRange(Helpers.InvokeScriptblockReturnStringArray(RowStyleScript, inputObject));

                int colspanCount = 0;
                resultBuilder.Append("<tr>\r\n");

                var cellClassDict = new Dictionary<string, string[]>(StringComparer.CurrentCultureIgnoreCase);
                if (CellClassScripts != null)
                {
                    foreach (DictionaryEntry pair in CellClassScripts)
                    {
                        string[] result = Helpers.InvokeScriptblockReturnStringArray(pair.Value, inputObject);
                        foreach (string header in Helpers.ConvertObjectToStringArray(pair.Key))
                            cellClassDict[header] = result;
                    }
                }

                var cellStyleDict = new Dictionary<string, string[]>(StringComparer.CurrentCultureIgnoreCase);
                if (CellStyleScripts != null)
                {
                    foreach (DictionaryEntry pair in CellStyleScripts)
                    {
                        string[] result = Helpers.InvokeScriptblockReturnStringArray(pair.Value, inputObject);
                        foreach (string header in Helpers.ConvertObjectToStringArray(pair.Key))
                            cellStyleDict[header] = result;
                    }
                }

                foreach (var header in headerList)
                {
                    bool skipCell = false;
                    if (colspanRowHash.ContainsKey(header) && colspanRowHash[header] > 0)
                    {
                        colspanRowHash[header] = colspanRowHash[header] - 1;
                        colspanCount = colspanCountHash[header];
                        skipCell = true;
                    }
                    if (colspanCount > 0)
                    {
                        colspanCount--;
                        skipCell = true;
                    }
                    if (rowspanCountHash.ContainsKey(header) && rowspanCountHash[header] > 0)
                    {
                        rowspanCountHash[header] = rowspanCountHash[header] - 1;
                        skipCell = true;
                    }

                    if (skipCell)
                        continue;

                    var cellClassList = new List<string>(rowClassList);
                    var cellStyleList = new List<string>(rowStyleList);

                    if (cellClassDict.ContainsKey(header))
                        cellClassList.AddRange(cellClassDict[header]);

                    if (cellStyleDict.ContainsKey(header))
                        cellStyleList.AddRange(cellStyleDict[header]);

                    if (lineClassDict.ContainsKey(header))
                        cellClassList.Add(string.Format("Insert{0}Line", lineClassDict[header]));

                    string cellValue = "";
                    if (inputObject.Properties[header] != null && inputObject.Properties[header].Value != null)
                        cellValue = String.Join(" ", Helpers.ConvertObjectToStringArray(inputObject.Properties[header].Value));

                    if (!HtmlProperty.Contains(header) && !(AutoDetectHtml.IsPresent && cellValue.Length > 0 && cellValue.Substring(0,1) == "<"))
                        cellValue = System.Web.HttpUtility.HtmlEncode(cellValue).Replace("\r\n", "<br />");

                    if (RightAlignProperty.Contains(header))
                        cellClassList.Add("ralign");

                    if (noWrapAll || NoWrapProperty.Contains(header))
                        cellClassList.Add("nowrap");

                    resultBuilder.Append("<td");

                    bool setColspan = false;
                    if (CellColspanScripts != null && CellColspanScripts.Contains(header))
                    {
                        colspanCount = GetCountFromProperty(inputObject, CellColspanScripts, header);
                        if (colspanCount > 1)
                        {
                            resultBuilder.AppendFormat(" colspan='{0}'", colspanCount);
                            colspanCount--;
                            setColspan = true;
                        }
                        else
                        {
                            colspanCount = 0;
                        }
                    }

                    int rowspanCount = 0;
                    bool setRowspan = false;
                    if (CellRowspanScripts != null && CellRowspanScripts.Contains(header))
                    {
                        rowspanCount = GetCountFromProperty(inputObject, CellRowspanScripts, header);
                        if (rowspanCount > 1)
                        {
                            resultBuilder.AppendFormat(" rowspan='{0}'", rowspanCount);
                            rowspanCountHash[header] = rowspanCount - 1;
                            setRowspan = true;
                        }
                    }

                    if (setRowspan && setColspan)
                    {
                        colspanRowHash[header] = rowspanCount - 1;
                        colspanCountHash[header] = colspanCount + 1;
                    }

                    if (cellClassList.Count > 0)
                        resultBuilder.AppendFormat(" class='{0}'", String.Join(" ", cellClassList));

                    if (cellStyleList.Count > 0)
                        resultBuilder.AppendFormat(" style='{0}'", String.Join(" ", cellStyleList));

                    if (AddDataColumnName.IsPresent)
                        resultBuilder.AppendFormat(" data-column-name='{0}'", System.Web.HttpUtility.HtmlAttributeEncode(header));

                    resultBuilder.AppendFormat(">{0}</td>\r\n", cellValue);
                }

                resultBuilder.Append("</tr>\r\n");
            }

            if (!RowsOnly.IsPresent && inputObjectList.Count != 0)
            {
                resultBuilder.Append("</tbody>\r\n");
                resultBuilder.Append("</table>");
            }

            WriteObject(resultBuilder.ToString());
        }
    }
}
