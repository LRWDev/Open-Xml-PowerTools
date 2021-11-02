using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OpenXmlPowerTools
{
    public class DocumentAssembler: DocumentAssemblerBase<XElement>
    {
        protected override string EvaluateExpression(object data, string xPath, bool optional)
        {

            var element = (XElement) data;
            object xPathSelectResult;
            try
            {
                //support some cells in the table may not have an xpath expression.
                if (String.IsNullOrWhiteSpace(xPath)) return String.Empty;

                xPathSelectResult = element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new XPathException("XPathException: " + e.Message, e);
            }

            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable)xPathSelectResult).Cast<XObject>();
                if (!selectedData.Any())
                {
                    if (optional) return string.Empty;
                    throw new XPathException(string.Format("XPath expression ({0}) returned no results", xPath));
                }
                if (selectedData.Count() > 1)
                {
                    throw new XPathException(string.Format("XPath expression ({0}) returned more than one node", xPath));
                }

                XObject selectedDatum = selectedData.First();

                if (selectedDatum is XElement) return ((XElement)selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute)selectedDatum).Value;
            }

            return xPathSelectResult.ToString();
        }

        protected override IEnumerable<object> SelectElements(object data, string expression)
        {
            var element = (XElement)data;
            return element.XPathSelectElements(expression);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError)
        {
            var docAsm = new DocumentAssembler();
            return docAsm.AssembleDocumentInternal(templateDoc, data, out templateError);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError)
        {
            var xDoc = data.GetXDocument();
            var docAsm = new DocumentAssembler();
            return docAsm.AssembleDocumentInternal(templateDoc, xDoc.Root, out templateError);
        }
    }
}
