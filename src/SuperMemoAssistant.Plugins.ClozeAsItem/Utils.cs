using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Runtime.Remoting;

namespace SuperMemoAssistant.Plugins.ClozeAsItem
{
  public static class Utils
  {

    public static IControlHtml GetFocusedHtmlCtrl()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        return ctrlGroup?.FocusedControl?.AsHtml();
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }

    /// <summary>
    /// Get the HTML string content representing the first html control of the current element.
    /// </summary>
    /// <returns>HTML string or null</returns>
    public static string GetCurrentElementContent()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.GetFirstHtmlControl()?.AsHtml();
        return htmlCtrl?.Text;
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }

    /// <summary>
    /// Get the selection object representing the currently highlighted text in SM.
    /// </summary>
    /// <returns>IHTMLTxtRange object or null</returns>
    public static IHTMLTxtRange GetSelectionObject()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        var sel = htmlDoc?.selection;

        if (!(sel?.createRange() is IHTMLTxtRange textSel))
          return null;

        return textSel;
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }
  }
}
