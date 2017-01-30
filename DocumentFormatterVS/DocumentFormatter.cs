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
            if (package.IsFormattingEnabled)
            {
                Document document = FindDocument(docCookie);
                if (!IsFormattingIgnored(document))
                {
                    FormatDocument(document);
                }
            }

            return VSConstants.S_OK;
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
    }
}
