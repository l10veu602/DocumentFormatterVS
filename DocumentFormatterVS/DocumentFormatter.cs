using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocumentFormatterVS
{
    class DocumentFormatter : IVsRunningDocTableEvents3
    {
        private DTE DTE;
        private RunningDocumentTable runningDocumentTable;
        private VSPackage1 package;

        public DocumentFormatter(DTE DTE, RunningDocumentTable runningDocumentTable, VSPackage1 package)
        {
            this.DTE = DTE;
            this.runningDocumentTable = runningDocumentTable;
            this.package = package;

            runningDocumentTable.Advise(this);
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeSave(uint docCookie)
        {
            int result = VSConstants.S_OK;
            if (package.IsFormattingDisabled)
            {
                return result;
            }

            Document document = FindDocument(docCookie);
            if (IsFormattingIgnored(document))
            {
                return result;
            }

            FormatDocument(document);

            if (package.IsUE4UPropertyFormattingEnabled && document.Name.ToLower().EndsWith(".h"))
            {
                var textDocument = document.Object("TextDocument") as TextDocument;
                FormatUProperty(textDocument);
            }

            return result;
        }

        private bool IsFormattingIgnored(Document document)
        {
            if (package.FormattingIgnoreRegexes == null)
            {
                return false;
            }

            foreach (var regex in package.FormattingIgnoreRegexes)
            {
                try
                {
                    if (Regex.IsMatch(document.FullName, regex))
                    {
                        return true;
                    }
                }
                catch { }
            }

            return false;
        }

        private Document FindDocument(uint docCookie)
        {
            var documentInfo = runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            return DTE.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentPath);
        }

        private void FormatDocument(Document document)
        {
            var currentDoc = DTE.ActiveDocument;

            document.Activate();

            if (DTE.ActiveWindow.Kind == "Document")
            {
                DTE.ExecuteCommand("Edit.FormatDocument");
            }

            currentDoc.Activate();
        }

        private void FormatUProperty(TextDocument textDocument)
        {
            if (textDocument == null)
            {
                return;
            }

            var editPoint = textDocument.CreateEditPoint();
            while (!editPoint.AtEndOfDocument)
            {
                string line = editPoint.GetText(editPoint.LineLength);
                editPoint.LineDown();

                if (editPoint.AtEndOfDocument)
                {
                    break;
                }

                var match = Regex.Match(line, @"^(\s*)((UPROPERTY)|(UFUNCTION))\(");
                if (!match.Success)
                {
                    continue;
                }

                string upropertyIndent = match.Groups[1].Value;
                var upropertyMemberText = editPoint.GetText(editPoint.LineLength);
                upropertyMemberText = upropertyIndent + upropertyMemberText.TrimStart(null);

                editPoint.ReplaceText(editPoint.LineLength, upropertyMemberText, 0);
                editPoint.LineDown();
            }
        }
    }
}
