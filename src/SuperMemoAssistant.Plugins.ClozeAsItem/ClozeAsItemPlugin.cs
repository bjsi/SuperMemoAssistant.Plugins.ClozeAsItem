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
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
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

    #endregion

    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      Svc.HotKeyManager.RegisterGlobal(
        "ClozeAsItem",
        "Create a new Cloze as an Item",
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
        if (parentEl == null || parentEl.Type == ElementType.Item)
          return;

        var references = ReferenceParser.GetReferences(content);
        if (references == null)
          return;

        int textSelStartIdx = Utils.GetSelectionTextStartIdx(selObj);
        int textSelEndIdx = Utils.GetSelectionTextEndIdx(selObj);

        if (textSelStartIdx == -1 || textSelEndIdx == -1 || textSelEndIdx < textSelStartIdx)
          return;

        string selHtml = selObj.htmlText;
        string selText2 = selObj.text;

        int htmlSelStartIdx = Utils.ConvertTextIdxToHtmlIdx(content, textSelStartIdx);

        // int htmlSelEndIdx = Utils.ConvertTextIdxToHtmlIdx(content, textSelEndIdx);

        string preCloze = content.Substring(0, htmlSelStartIdx);

        string postCloze = content.Substring(htmlSelStartIdx + selHtml.Length);

        string question = preCloze + "<span class='cloze'>[...]</span>" + postCloze;
        string answer = content.Substring(htmlSelStartIdx, selHtml.Length);
        if (string.IsNullOrEmpty(question) || string.IsNullOrEmpty(answer))
          return;

        CreateSMElement(question, answer, parentEl);

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
        LogTo.Error("Failed to CreateSMElement because parent element was null");
        return;
      }

      if (parent.Id != Svc.SM.UI.ElementWdw.CurrentElementId)
      {
        LogTo.Debug("Failed to CreateSMElement because the displayed element changed");
        return;
      }

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var ctrl = ctrlGroup[i];
        if (ctrl == null || i == ctrlGroup.FocusedControlIndex)
          continue;

        // TODO: How to maintain same size, layout options
        // TODO: Add other component types
        // TODO: Create a generic 'inherit parent components' utility

        switch (ctrl.Type)
        {

          case ComponentType.Image:
            var image = ctrl as IControlImage;
            contents.Add(new ImageContent(image.ImageMemberId));
            break;

          default:
            break;

        }

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

    // Set HasSettings to true, and uncomment this method to add your custom logic for settings
    // /// <inheritdoc />
    // public override void ShowSettings()
    // {
    // }

    #endregion

    #region Methods

    #endregion
  }
}
