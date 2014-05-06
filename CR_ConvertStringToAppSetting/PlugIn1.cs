using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.CodeRush.Core;
using DevExpress.CodeRush.PlugInCore;
using DevExpress.CodeRush.StructuralParser;
using SP = DevExpress.CodeRush.StructuralParser;
using DevExpress.Refactor;

namespace CR_ConvertStringToAppSetting
{
    public partial class PlugIn1 : StandardPlugIn
    {
        // DXCore-generated code...
        #region InitializePlugIn
        public override void InitializePlugIn()
        {
            base.InitializePlugIn();
            registerConvertStringToAppSetting();
        }
        #endregion
        #region FinalizePlugIn
        public override void FinalizePlugIn()
        {
            base.FinalizePlugIn();
        }
        #endregion


        public void registerConvertStringToAppSetting()
        {
            DevExpress.Refactor.Core.RefactoringProvider ConvertStringToAppSetting = new DevExpress.Refactor.Core.RefactoringProvider(components);
            ((System.ComponentModel.ISupportInitialize)(ConvertStringToAppSetting)).BeginInit();
            ConvertStringToAppSetting.ProviderName = "ConvertStringToAppSetting"; // Should be Unique
            ConvertStringToAppSetting.DisplayName = "Convert String to AppSetting";
            ConvertStringToAppSetting.CheckAvailability += ConvertStringToAppSetting_CheckAvailability;
            ConvertStringToAppSetting.Apply += ConvertStringToAppSetting_Execute;
            ((System.ComponentModel.ISupportInitialize)(ConvertStringToAppSetting)).EndInit();
        }
        private void ConvertStringToAppSetting_CheckAvailability(Object sender, CheckContentAvailabilityEventArgs ea)
        {
            PrimitiveExpression Literal = ea.Element as PrimitiveExpression;
            if (Literal == null)
                return;
            if (Literal.PrimitiveType != PrimitiveType.String)
                return;
            ea.Available = true;
        }
        private void ConvertStringToAppSetting_Execute(Object sender, ApplyContentEventArgs ea)
        {
            TextDocument CodeDoc = CodeRush.Documents.ActiveTextDocument;
            using (CodeDoc.NewCompoundAction("Extract to App.config setting"))
            {
                ProjectElement ActiveProject = CodeRush.Source.ActiveProject;
                PrimitiveExpression StringLiteral = ea.Element as PrimitiveExpression;

                // Ensure App.config 
                string BasePath = System.IO.Path.GetDirectoryName(ActiveProject.FilePath);
                string Filename = BasePath + "\\App.Config";
                TextDocument configDoc = CodeRush.Documents.GetTextDocument(Filename);
                
                if (configDoc == null) 
                {
                    // Open configDoc
                    configDoc = (TextDocument)CodeRush.File.Activate(Filename);
                }

                if (configDoc == null) 
                {
                    // Create configDoc
                    string DefaultSettings = "<?xml version=\"1.0\" encoding=\"utf-8\" ?><configuration></configuration>";
                    CodeRush.File.WriteTextFile(Filename, DefaultSettings);
                    CodeRush.Project.AddFile(ActiveProject, Filename);
                    configDoc = (TextDocument)CodeRush.File.Activate(Filename);
                }

                //Ensure RootNode
                SP.HtmlElement RootNode = (SP.HtmlElement)configDoc.FileNode.Nodes[1];

                // Get appSettings node
                SP.HtmlElement AppSettings = GetAppSettings(RootNode);
                if (AppSettings == null)
                {
                    AppSettings = CreateHTMLNode("appSettings");
                    RootNode.AddNode(AppSettings);
                    RewriteNodeInDoc(RootNode, configDoc);
                    RootNode = (SP.HtmlElement)configDoc.FileNode.Nodes[1];
                    AppSettings = GetAppSettings(RootNode);
                }

                // Generate a new setting... Add it to correct location in App.config.
                string SettingValue = (string)StringLiteral.PrimitiveValue;
                string SettingName = "MySetting";
                SP.HtmlElement SettingNode = CreateSettingNode(SettingName, SettingValue);
                AppSettings.AddNode(SettingNode);
                RewriteNodeInDoc(AppSettings, configDoc);

                // Add reference to System.Configuration dll.
                CodeRush.Project.AddReference(ActiveProject, "System.Configuration.dll");

                // Replace Literal with reference to setting through Configuration manager
                string Code = String.Format("System.Configuration.ConfigurationManager.AppSettings[\"{0}\"]", SettingName);
                var NewCodeRange = CodeDoc.SetText(StringLiteral.Range, Code);
                CodeDoc.ParseIfTextChanged();
                CodeDoc.ParseIfNeeded();
                var SourceNode = CodeDoc.GetNodeAt(NewCodeRange.Start);
                
                // Find Newly created Setting
                configDoc.ParseIfTextChanged();
                configDoc.ParseIfNeeded();
                RootNode = (SP.HtmlElement)configDoc.FileNode.Nodes[1];
                AppSettings = GetAppSettings(RootNode);
                SettingNode = GetSettingWithKeyAndValue(AppSettings, SettingName, SettingValue);

                // Link Code and Setting.
                var LinkSet = CodeRush.LinkedIdentifiers.NewMultiDocumentContainer();
                SourceRange StringRange = (SourceNode.Parent.Parent.Parent.Parent.DetailNodes[0] as LanguageElement).Range;
                SourceRange CodeSourceRange = new SourceRange(StringRange.Start.OffsetPoint(0, 1), StringRange.End.OffsetPoint(0, -1));
                LinkSet.Add(new FileSourceRange(CodeDoc.FileNode, CodeSourceRange));
                LinkSet.Add(new FileSourceRange(configDoc.FileNode, SettingNode.Attributes["key"].ValueRange));
                CodeRush.LinkedIdentifiers.Invalidate(configDoc);
                CodeDoc.Activate();
                CodeRush.LinkedIdentifiers.Invalidate(CodeDoc);
                CodeRush.Selection.SelectRange(CodeSourceRange);
                configDoc.ParseIfTextChanged();
                configDoc.ParseIfNeeded();

            }
        }

        private SP.HtmlElement GetSettingWithKeyAndValue(SP.HtmlElement appSettings, string settingName, string settingValue)
        {
            return (from SP.HtmlElement item in appSettings.Nodes
                    where item.Attributes["value"].Value == settingValue
                    && item.Attributes["key"].Value == settingName
                    select item).First();
        }
        private SP.HtmlElement GetAppSettings(SP.HtmlElement RootNode)
        {
            var TopLevelNodes = from SP.HtmlElement item in RootNode.Nodes select item;
            var AppSettings = (from SP.HtmlElement item in TopLevelNodes
                               where item.Name == "appSettings"
                               select item).FirstOrDefault();
            return AppSettings;
        }
        private void RewriteNodeInDoc(LanguageElement Node, TextDocument Doc)
        {
            var Code = CodeRush.Language.GenerateElement(Node, Doc.Language);
            Doc.SetText(Node.Range, Code);
        }
        private static SP.HtmlElement CreateHTMLNode(string NodeName, bool EmptyTag = false)
        {
            var Node = new SP.HtmlElement();
            Node.Name = NodeName;
            Node.IsEmptyTag = EmptyTag;
            return Node;
        }
        private static SP.HtmlElement CreateSettingNode(string SettingName, string SettingValue)
        {
            var setting = CreateHTMLNode("add", true);
            setting.AddAttribute("key", SettingName);
            setting.AddAttribute("value", SettingValue);
            return setting;
        }


    }
    public class StringPrimitiveFilter : IElementFilter
    {
        public bool Apply(IElement element)
        {
            if (element.ElementType != LanguageElementType.PrimitiveExpression)
                return false;
            PrimitiveExpression Expression = (PrimitiveExpression)element;
            if (Expression.PrimitiveType != PrimitiveType.String)
                return false;
            return true;
        }

        public bool SkipChildren(IElement element)
        {
            return false;
        }
    }
}