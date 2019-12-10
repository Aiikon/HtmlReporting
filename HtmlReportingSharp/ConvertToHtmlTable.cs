﻿using System;
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
        public string[] HtmlProperty { get; set; }

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
        public string[] RightAlignProperty { get; set; }

        [Parameter()]
        public string[] NoWrapProperty { get; set; }

        [Parameter()]
        public SwitchParameter Narrow { get; set; }

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

            if (Property == null)
                Property = inputObjectList[0].Properties.Select(p => p.Name).ToArray();

            foreach (var p in Property)
                headerList.Add(p);

            if (!RowsOnly.IsPresent)
            {
                resultBuilder.Append("<table class='HtmlReportingTable");
                if (Narrow.IsPresent)
                    resultBuilder.Append(" Narrow");
                resultBuilder.Append("'>\r\n");
                resultBuilder.Append("<thead>\r\n");
                resultBuilder.Append("<tr class='header'>\r\n");
                foreach (var header in headerList)
                    resultBuilder.Append(String.Format("<th>{0}</th>\r\n", header));
                resultBuilder.Append("</tr>\r\n");
                resultBuilder.Append("</thead>\r\n");
                resultBuilder.Append("<tbody>\r\n");
            }

            var rowspanCountHash = new Dictionary<string, int>();
            int rowspanCount = 0;

            foreach (var inputObject in inputObjectList)
            {
                var rowClassList = new List<string>();
                var rowStyleList = new List<string>();

                if (RowClassScript != null)
                    foreach (var result in RowClassScript.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", inputObject) }, null))
                        if (result != null)
                            rowClassList.Add(result.ToString());

                if (RowStyleScript != null)
                    foreach (var result in RowStyleScript.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", inputObject) }, null))
                        if (result != null)
                            rowStyleList.Add(result.ToString());

                int colspanCount = 0;
                resultBuilder.Append("<tr>\r\n");

                foreach (var header in headerList)
                {
                    bool skipCell = false;
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

                    string cellValue = "";
                    if (inputObject.Properties[header] != null && inputObject.Properties[header].Value != null)
                        cellValue = inputObject.Properties[header].Value.ToString();

                    if (CellClassScripts != null && CellClassScripts.ContainsKey(header))
                    {
                        ScriptBlock script = (ScriptBlock)CellClassScripts[header];
                        foreach (var result in script.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", inputObject) }, null))
                            if (result != null)
                                cellClassList.Add(result.ToString());
                    }

                    if (CellStyleScripts != null && CellStyleScripts.ContainsKey(header))
                    {
                        ScriptBlock script = (ScriptBlock)CellStyleScripts[header];
                        foreach (var result in script.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("_", inputObject) }, null))
                            if (result != null)
                                cellStyleList.Add(result.ToString());
                    }

                    if (!HtmlProperty.Contains(header))
                        cellValue = System.Web.HttpUtility.HtmlEncode(cellValue).Replace("\r\n", "<br />");

                    if (RightAlignProperty.Contains(header))
                        cellClassList.Add("ralign");

                    if (noWrapAll || NoWrapProperty.Contains(header))
                        cellClassList.Add("nowrap");

                    resultBuilder.Append("<td");

                    if (CellColspanScripts != null && CellColspanScripts.Contains(header))
                    {
                        colspanCount = GetCountFromProperty(inputObject, CellColspanScripts, header);
                        if (colspanCount > 1)
                        {
                            resultBuilder.AppendFormat(" colspan='{0}'", colspanCount);
                            colspanCount--;
                        }
                    }

                    if (CellRowspanScripts != null && CellRowspanScripts.Contains(header))
                    {
                        rowspanCount = GetCountFromProperty(inputObject, CellRowspanScripts, header);
                        if (rowspanCount > 1)
                        {
                            resultBuilder.AppendFormat(" rowspan='{0}'", rowspanCount);
                            rowspanCount--;
                        }
                    }

                    if (cellClassList.Count > 0)
                        resultBuilder.AppendFormat(" class='{0}'", String.Join(" ", cellClassList));

                    if (cellStyleList.Count > 0)
                        resultBuilder.AppendFormat(" style='{0}'", String.Join(" ", cellStyleList));

                    resultBuilder.AppendFormat(">{0}</td>\r\n", cellValue);
                }

                resultBuilder.Append("</tr>\r\n");
            }

            if (!RowsOnly.IsPresent)
            {
                resultBuilder.Append("</tbody>\r\n");
                resultBuilder.Append("</table>");
            }

            WriteObject(resultBuilder.ToString());
        }
    }
}