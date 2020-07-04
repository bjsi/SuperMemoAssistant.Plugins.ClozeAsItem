using HtmlAgilityPack;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SuperMemoAssistant.Plugins.ClozeAsItem
{
  public static class Utils
  {

    public static void TestingSelectionObj(IHTMLTxtRange selObj, int index)
    {

      while (index > 0)
      {

        selObj.moveStart("character", -1);
        selObj.moveEnd("character", -1);
        index -= 1;
        string text = selObj.text;
        string htmlText = selObj.htmlText;

      }

    }

    public static IControlHtml GetFocusedCtrl()
    {
      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      return htmlCtrl;
    }

    /// <summary>
    /// Get the HTML string content representing the first html control of the current element.
    /// </summary>
    /// <returns>HTML string or null</returns>
    public static string GetCurrentElementContent()
    {
      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.GetFirstHtmlControl()?.AsHtml();
      return htmlCtrl?.Text;
    }

    /// <summary>
    /// Get the selection object representing the currently highlighted text in SM.
    /// </summary>
    /// <returns>IHTMLTxtRange object or null</returns>
    public static IHTMLTxtRange GetSelectionObject()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      var htmlDoc = htmlCtrl?.GetDocument();
      var sel = htmlDoc?.selection;

      if (!(sel?.createRange() is IHTMLTxtRange textSel))
        return null;

      return textSel;

    }

    /// <summary>
    /// Get the end index of the selection object as an inner text index.
    /// This is a "one past the end" index. A selection would be up to and NOT including this index.
    /// </summary>
    /// <param name="selObj"></param>
    /// <returns>index or -1</returns>
    public static int GetSelectionTextEndIdx(IHTMLTxtRange selObj)
    {

      int MaxTextLength = 2000000000;
      int result = -1;
      if (selObj != null)
      {
        var duplicate = selObj.duplicate();
        result = Math.Abs(duplicate.moveEnd("character", -MaxTextLength));
      }

      return result;

    }

    /// <summary>
    /// Get the start index of the selection object as an inner text index.
    /// </summary>
    /// <param name="selObj"></param>
    /// <returns>index or -1</returns>
    public static int GetSelectionTextStartIdx(IHTMLTxtRange selObj)
    {

      int MaxTextLength = 2000000000;
      int result = -1;
      if (selObj != null)
      {
        var duplicate = selObj.duplicate();
        result = Math.Abs(duplicate.move("character", -MaxTextLength));
      }
      return result;

    }

    /// <summary>
    /// Convert a text index to the equivalent position in the html string.
    /// </summary>
    /// <param name="html"></param>
    /// <param name="textIdx"></param>
    /// <returns>index or -1</returns>
    public static int ConvertTextIdxToHtmlIdx(string html, int textIdx)
    {

      if (string.IsNullOrEmpty(html))
        return -1;

      if (textIdx < 0)
        return -1;

      var doc = new HtmlDocument();
      doc.LoadHtml(html);

      var nodes = doc.DocumentNode
                    ?.Descendants()
                    ?.Where(x => x.Name == "#text" || x.Name == "br");

      if (nodes == null || !nodes.Any())
        return -1;

      // Return -1 if not found
      int htmlIdx = -1;

      int startIdx = -1;
      int endIdx = -1;


      // Last pair of characters
      Tuple<char, char> last = new Tuple<char, char>('x', 'x');

      foreach (var node in nodes)
      {

        if (node.Name == "br")
        {
          startIdx++;
          endIdx++;
          continue;
        }

        var decoded = HttpUtility.HtmlDecode(node.InnerText);

        // Consecutive \r\n get turned into a single null object in the IHTMLTxtRange
        // Increment TextIdx for each consecutive \r\n

        if (decoded.Length == 1)
        {
          last = new Tuple<char, char>(last.Item2, decoded[0]);
        }
        else
        {
          for (int i = 0, j = 1; j < decoded.Length; i++, j++)
          {
            if (decoded[i] == '\r' && decoded[j] == '\n')
            {
              if (last.Item1 == '\r' && last.Item2 == '\n')
              {
                textIdx++;
              }
            }
            last = new Tuple<char, char>(decoded[i], decoded[j]);
          }
        }

        var length = decoded.Replace("\r\n", " ").Length;
        endIdx += length;

        if (startIdx <= textIdx && textIdx <= endIdx)
        {
          if (startIdx == -1)
            startIdx = 0;

          var textIdxDiff = (textIdx - startIdx);
          var htmlIdxDiff = 0;

          // Add the length of each html encoded character before the target
          // to the htmlIdx
          // TODO: This encodes punctuation like ' 
          // TODO: Check which characters I need to manage

          //for (int i = 0; i < textIdxDiff; i++)
          //  htmlIdxDiff += HttpUtility.HtmlEncode(decoded[i].ToString()).Length;

          htmlIdx = node.InnerStartIndex + textIdxDiff;
          break;
        }

        // TODO: Testing moving here
        if (endIdx > -1)
          startIdx = endIdx + 1;
      }

      return htmlIdx;
    }
  }
}
