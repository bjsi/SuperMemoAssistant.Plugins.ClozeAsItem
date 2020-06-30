#region License & Metadata

// The MIT License (MIT)
// 
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the 
// Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
// 
// 
// Created On:   6/30/2020 7:39:54 PM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.ClozeAsItem
{
  using System;
  using System.Collections.Generic;
  using System.Diagnostics.CodeAnalysis;
  using System.Runtime.Remoting;
  using System.Windows.Input;
  using Anotar.Serilog;
  using HtmlAgilityPack;
  using mshtml;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Models;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Types;
  using SuperMemoAssistant.Services;
  using SuperMemoAssistant.Services.IO.Keyboard;
  using SuperMemoAssistant.Services.Sentry;
  using SuperMemoAssistant.Sys.IO.Devices;

  // ReSharper disable once UnusedMember.Global
  // ReSharper disable once ClassNeverInstantiated.Global
  [SuppressMessage("Microsoft.Naming", "CA1724:TypeNamesShouldNotMatchNamespaces")]
  public class ClozeAsItemPlugin : SentrySMAPluginBase<ClozeAsItemPlugin>
  {
    #region Constructors

    /// <inheritdoc />
    public ClozeAsItemPlugin() : base("Enter your Sentry.io api key (strongly recommended)") { }

    #endregion


    #region Properties Impl - Public

    /// <inheritdoc />
    public override string Name => "ClozeAsItem";

    /// <inheritdoc />
    public override bool HasSettings => false;
    private const int MaxTextLength = 2000000000;

    #endregion


    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      Svc.HotKeyManager.RegisterGlobal(
        "ItemCloze",
        "Create new Item Cloze",
        HotKeyScopes.SMBrowser,
        new HotKey(Key.Z, KeyModifiers.CtrlAltShift),
        CreateItemCloze
      );

    }

    [LogToErrorOnException]
    private void CreateItemCloze()
    {

      try
      {

        var selObj = Utils.GetSelectionObject();
        string selText = selObj?.text;
        if (selObj == null || string.IsNullOrEmpty(selText))
          return;

        var content = Utils.GetCurrentElementContent();
        if (string.IsNullOrEmpty(content))
          return;

        var parentEl = Svc.SM.UI.ElementWdw.CurrentElement;
        if (parentEl == null)
          return;

        var references = ReferenceParser.GetReferences(content);
        if (references == null)
          return;

        int textSelStart = GetSelStart(selObj);
        int textSelEnd = GetSelEnd(selObj);
        if (textSelStart == -1 || textSelEnd == -1 || textSelEnd < textSelStart)
          return;

        int htmlSelStart = ConvertTextIdxToHtmlIdx(content, textSelStart);
        int htmlSelEnd = ConvertTextIdxToHtmlIdx(content, textSelEnd);
        int htmlSelLength = htmlSelEnd - htmlSelStart;

        string question = content.Substring(0, htmlSelStart) + "<span class='cloze'>[...]</span>" + content.Substring(htmlSelEnd);
        string answer = content.Substring(htmlSelStart, htmlSelLength);
        if (string.IsNullOrEmpty(question) || string.IsNullOrEmpty(answer))
          return;

        //CreateSMElement(question, answer, parentEl);

      }
      catch (RemotingException) { }

    }

    [LogToErrorOnException]
    private void CreateSMElement(string question, string answer, IElement parent)
    {

      var contents = new List<ContentBase>();

      contents.Add(new TextContent(true, question));
      contents.Add(new TextContent(true, answer, displayAt: AtFlags.NonQuestion));

      if (parent == null)
      {
        LogTo.Error("Failed to CreateSMElement beacuse parent element was null");
        return;
      }

      bool success = Svc.SM.Registry.Element.Add(
        out var value,
        ElemCreationFlags.ForceCreate,
        new ElementBuilder(ElementType.Item, contents.ToArray())
          .WithParent(parent)
          .WithLayout("Item")
          .WithPriority(30)
          .DoNotDisplay()
      );

      if (success)
      {
        LogTo.Debug("Successfully created SM Element");
      }
      else
      {
        LogTo.Error("Failed to CreateSMElement");
      }
    }

    // TODO: Not Working
    private int ConvertTextIdxToHtmlIdx(string html, int textIdx)
    {
      var doc = new HtmlDocument();
      doc.LoadHtml(html);
      var nodes = doc.DocumentNode.Descendants();

      int startIdx = 0;
      int endIdx = 0;

      int htmlIdx = -1;

      foreach (var node in nodes)
      {
        endIdx += node.InnerText?.Length ?? 0;
        if (startIdx <= textIdx && textIdx <= endIdx)
        {
          htmlIdx = node.InnerStartIndex + (textIdx - startIdx);
          break;
        }
      }

      return htmlIdx;
    }

    private int GetSelEnd(IHTMLTxtRange selObj)
    {
      int result = -1;
      if (selObj != null)
      {
        var duplicate = selObj.duplicate();
        result = Math.Abs(duplicate.moveEnd("character", -MaxTextLength));
      }
      return result;
    }

    private int GetSelStart(IHTMLTxtRange selObj)
    {
      int result = -1;
      if (selObj != null)
      {
        var duplicate = selObj.duplicate();
        result = Math.Abs(duplicate.moveStart("character", -MaxTextLength));
      }
      return result;
    }

    // Set HasSettings to true, and uncomment this method to add your custom logic for settings
    // /// <inheritdoc />
    // public override void ShowSettings()
    // {
    // }

    #endregion




    #region Methods

    // Uncomment to register an event handler for element changed events
    // [LogToErrorOnException]
    // public void OnElementChanged(SMDisplayedElementChangedEventArgs e)
    // {
    //   try
    //   {
    //     Insert your logic here
    //   }
    //   catch (RemotingException) { }
    // }

    #endregion
  }
}
