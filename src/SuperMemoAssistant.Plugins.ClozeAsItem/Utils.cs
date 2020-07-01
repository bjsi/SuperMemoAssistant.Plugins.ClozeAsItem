using HtmlAgilityPack;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.ClozeAsItem
{
  public static class Utils
  {

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

      // TODO: Testing
      return result - 1;

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
        result = Math.Abs(duplicate.moveStart("character", -MaxTextLength));
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
                    ?.Where(x => x.Name == "#text");
      if (nodes == null || !nodes.Any())
        return -1;

      // Return -1 if not found
      int htmlIdx = -1;

      int startIdx = 0;
      // Starts on -1 because we are adding 1-indexed length to 0-indexed index
      int endIdx = -1;

      foreach (var node in nodes)
      {

        endIdx += node.InnerText.Length;

        if (startIdx <= textIdx && textIdx <= endIdx)
        {
          htmlIdx = node.InnerStartIndex + (textIdx - startIdx);
          break;
        }

        startIdx = endIdx + 1;
      }

      return htmlIdx;
    }
  }
}
