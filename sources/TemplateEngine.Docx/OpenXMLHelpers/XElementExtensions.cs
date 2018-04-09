using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;

namespace TemplateEngine.Docx
{
	static class XElementExtensions
	{
        public static string GetUnicodeString(string code)
        {
            if (code == null) return null;
            var bytes = new byte[code.Length / 2];
            for(var i = 0; i < bytes.Length ; i++)
            {
                var str = code.Substring(i*2, 2);
                bytes[bytes.Length - i - 1] = byte.Parse(str, System.Globalization.NumberStyles.AllowHexSpecifier);
            }
            return Encoding.Unicode.GetString(bytes);
        }

		// Set content control value th the new value
		public static void ReplaceContentControlWithNewValue(this XElement sdt, string newValue)
		{
            // We need to check if the field is a checkbox
            var sdtPr = sdt.Element(W.sdtPr);
            bool isCheckbox = false;
            bool isDate = false;
            string chkValueIfTrue = "";
            string chkValueIfFalse = "";
            string chkFontIfTrue = "";
            string chkFontIfFalse = "";
            string dateFormat = "";
            if (sdtPr != null)
            {
                var checkbox = sdtPr.Element(W14.checkbox);
                if (checkbox != null)
                {
                    // Checkbox! Extract the correct values for true and false
                    isCheckbox = true;

                    var checkedState = checkbox.Element(W14.checkedState);
                    chkValueIfTrue = GetUnicodeString(checkedState.Attribute(W14.val)?.Value);
                    chkFontIfTrue = checkedState.Attribute(W14.font)?.Value;
                    if (chkValueIfTrue == null)
                    { // https://msdn.microsoft.com/en-us/library/dd947495(v=office.12).aspx
                        chkValueIfTrue = GetUnicodeString("2612");
                        chkFontIfTrue = "MS Gothic";
                    }

                    var uncheckedState = checkbox.Element(W14.uncheckedState);
                    chkValueIfFalse = GetUnicodeString(uncheckedState.Attribute(W14.val)?.Value);
                    chkFontIfFalse = uncheckedState.Attribute(W14.font)?.Value;
                    if (chkValueIfFalse == null)
                    { // https://msdn.microsoft.com/en-us/library/dd947495(v=office.12).aspx
                        chkValueIfTrue = GetUnicodeString("2610");
                        chkFontIfTrue = "MS Gothic";
                    }
                } else
                {
                    var date = sdtPr.Element(W.date);
                    if (date != null)
                    {
                        dateFormat = date.Element(W.dateFormat)?.Attribute(W.val)?.Value;
                        isDate = true;
                    }
                }
            }

			var sdtContentElement = sdt.Element(W.sdtContent);

			if (sdtContentElement != null)
			{
				var elementsWithText = sdtContentElement.Elements()
					.Where(e => 
						e.DescendantsAndSelf(W.t).Any() &&
						!e.DescendantsAndSelf(W.sdt).Any())						
					.ToList();
				
				var firstContentElementWithText = elementsWithText.FirstOrDefault(d => d.DescendantsAndSelf(W.t).Any());

				if (firstContentElementWithText != null)
				{
					var firstTextElement = firstContentElementWithText
						.Descendants(W.t)
						.First();

                    if (isCheckbox)
                    {
                        // Very dirty my dear. Nasty boy!
                        var boolValue = bool.Parse(newValue);
                        firstTextElement.Value = boolValue ? chkValueIfTrue : chkValueIfFalse;
                        var pr = firstTextElement.ElementsBeforeSelf().FirstOrDefault(d => d.DescendantsAndSelf(W.rFonts).Any());
                        if (pr != null)
                        {
                            var font = pr.DescendantsAndSelf(W.rFonts).FirstOrDefault();
                            font.Attribute(W.ascii).Value = boolValue ? chkFontIfTrue : chkFontIfFalse;
                            font.Attribute(W.eastAsia).Value = boolValue ? chkFontIfTrue : chkFontIfFalse;
                            font.Attribute(W.hAnsi).Value = boolValue ? chkFontIfTrue : chkFontIfFalse;
                        }
                    }
                    else if (isDate)
                    {
                        firstTextElement.Value = DateTime.FromBinary(long.Parse(newValue)).ToString(dateFormat);
                    }
                    else
                    {
                        firstTextElement.Value = newValue;
                    }

                    //remove all text elements with its ancestors from the first contentElement
                    var firstElementAncestors = firstTextElement.AncestorsAndSelf().ToList();
					
					foreach (var descendants in elementsWithText.DescendantsAndSelf().ToList())
					{
						if (!firstElementAncestors.Contains(descendants) && descendants.DescendantsAndSelf(W.t).Any())
						{
							descendants.Remove();
						}
						//descendants.AncestorsAndSelf().Where(a => !firstElementAncestors.Contains(a)).Remove();
					}

					var contentReplacementElement = new XElement(firstContentElementWithText);
					
				/*	sdtContentElement.Descendants().Where(d => d.Descendants(W.t).Any() && d != firstContentElementWithText && d.Name != W.sdt).Remove();*/
					firstContentElementWithText.AddAfterSelf(contentReplacementElement);
					firstContentElementWithText.Remove();
				}
				else
				{
					if (sdtContentElement.Elements(W.p).Any())
					{
						sdtContentElement.Element(W.p).Add(new XElement(W.r, new XElement(W.t, newValue)));
					}
					else
					{
						sdtContentElement.Add(new XElement(W.p), new XElement(W.r, new XElement(W.t, newValue)));
					}
				}
			}
			else
			{
				sdt.Add(new XElement(W.sdtContent, new XElement(W.p), new XElement(W.r, new XElement(W.t, newValue))));
			}

			ReplaceNewLinesWithBreaks(sdt);
		}

		public static void RemoveContentControl(this XElement sdt)
		{

			var sdtContentElement = sdt.Element(W.sdtContent);
			if (sdtContentElement == null)
			{
				sdt.Remove();
				return;
			}

			var parent = new XElement("parent");
			if (sdt.Parent == null)
			{
				//add newElement to fake parent for remove content control
				parent.Add(sdt);
			}
			// Remove the content control, and replace it with its contents.
			sdt.ReplaceWith(sdtContentElement.Elements());

			if (sdt.Parent == parent)
			{
				sdt.Remove();
			}
		}
		public static IEnumerable<XElement> FirstLevelDescendantsAndSelf(this XElement element, XName name)
		{
			var allDescendantsAndSelf = element
				//content controls
				.DescendantsAndSelf(name).ToList();

			foreach (var xElement in allDescendantsAndSelf)
			{
				var ancestors = xElement.Ancestors(name);
				var inter = ancestors.Intersect(allDescendantsAndSelf);
				var count = inter.Count();
			}
			return allDescendantsAndSelf
				.Where(d => !(d.Ancestors(name).Intersect(allDescendantsAndSelf)).Any());
		}

		public static IEnumerable<XElement> FirstLevelDescendantsAndSelf(this IEnumerable<XElement> element, XName name)
		{
			var allDescendantsAndSelf = element
				//content controls
				.DescendantsAndSelf(name).ToList();


			return allDescendantsAndSelf
				.Where(d => !d.Ancestors().Any(allDescendantsAndSelf.Contains));
		}


		public static string SdtTagName(this XElement sdt)
		{
			if (sdt.Name != W.sdt) return null;

			try
			{
				return sdt
					.Element(W.sdtPr)
					.Element(W.tag)
					.Attribute(W.val)
					.Value;
			}
			catch (Exception)
			{
				return null;
			}
		}

		public static void ReplaceNewLinesWithBreaks(XElement xElem)
		{
			if (xElem == null) return;

			var textWithBreaks = xElem.Descendants(W.t).Where(t => t.Value.Contains("\r\n"));
			foreach (var textWithBreak in textWithBreaks)
			{
				var text = textWithBreak.Value;
				var split = text.Replace("\r\n", "\n").Split(new[] { "\n" }, StringSplitOptions.None);
				textWithBreak.Value = string.Empty;
				foreach (var s in split)
				{
					textWithBreak.Add(new XElement(W.t, s));
					textWithBreak.Add(new XElement(W.br));
				}
				textWithBreak.Descendants(W.br).Last().Remove();
			}
		}
	}
}
