using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Remoting;
using System.Threading.Tasks;
using System.Windows.Input;
using Anotar.Serilog;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
using SuperMemoAssistant.Interop.SuperMemo.Content.Models;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Services.IO.HotKeys;
using SuperMemoAssistant.Services.IO.Keyboard;
using SuperMemoAssistant.Services.Sentry;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.IO.Devices;

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
    public override bool HasSettings => true;

    public ClozeAsItemCfg Config { get; private set; }

    #endregion

    #region Methods Impl

    private async Task LoadConfig()
    {
      Config = await Svc.Configuration.Load<ClozeAsItemCfg>().ConfigureAwait(false) ?? new ClozeAsItemCfg();
    }


    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig().Wait();

      Svc.HotKeyManager.RegisterGlobal(
        "ClozeAsItem",
        "Create a new Cloze as an Item",
        HotKeyScope.SMBrowser,
        new HotKey(Key.Z, KeyModifiers.CtrlAltShift),
        CreateItemCloze
      );

    }

    [LogToErrorOnException]
    private void CreateItemCloze()
    {
      try
      {
        var parentEl = Svc.SM.UI.ElementWdw.CurrentElement;
        if (parentEl == null || parentEl.Type == ElementType.Item)
          return;

        var selObj = Utils.GetSelectionObject();
        string selText = selObj?.htmlText;
        if (selObj == null || string.IsNullOrEmpty(selText))
          return;

        var htmlCtrl = Utils.GetFocusedHtmlCtrl();
        var htmlDoc = htmlCtrl?.GetDocument();
        if (htmlDoc == null)
          return;
        
        selObj.pasteHTML("[...]");
        string questionChild = htmlDoc.body.innerHTML.Replace("[...]", "<SPAN class=cloze>[...]</SPAN>");

        int MaxTextLength = 2000000000;
        selObj.moveEnd("character", MaxTextLength);
        selObj.moveStart("character", -MaxTextLength);

        selObj.findText("[...]");
        selObj.select();

        selObj.pasteHTML("<SPAN class=clozed>" + selText + "</SPAN>");
        string parentContent = htmlDoc.body.innerHTML;
        
        htmlCtrl.Text = parentContent;

        var references = ReferenceParser.GetReferences(parentContent);
        CreateSMElement(RemoveReferences(questionChild), selText, references);

      }
      catch (RemotingException) { }
    }

    private string RemoveReferences(string htmlText)
    {
      var idx = htmlText.IndexOf("HR SuperMemo", System.StringComparison.InvariantCultureIgnoreCase);
      if (idx >= 0)
        return htmlText.Substring(0, idx + 1);
      return htmlText;
    }

    [LogToErrorOnException]
    private void CreateSMElement(string question, string answer, References refs)
    {

      var contents = new List<ContentBase>();
      var parent = Svc.SM.UI.ElementWdw.CurrentElement;
      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;

      contents.Add(new TextContent(true, question));
      contents.Add(new TextContent(true, answer, displayAt: AtFlags.NonQuestion));

      if (parent == null)
      {
        LogTo.Error("Failed to CreateSMElement because parent element was null");
        return;
      }

      bool success = Svc.SM.Registry.Element.Add(
        out _,
        ElemCreationFlags.ForceCreate,
        new ElementBuilder(ElementType.Item, contents.ToArray())
          .WithParent(parent)
          .WithLayout("Item")
          .WithPriority(30)
          .DoNotDisplay()
          .WithReference((_) => refs)
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
     /// <inheritdoc />
    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate(HotKeyManager.Instance, Config);
    }

    #endregion

    #region Methods
    #endregion
  }
}
