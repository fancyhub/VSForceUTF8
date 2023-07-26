using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.ComponentModel;
using Task = System.Threading.Tasks.Task;

namespace VsForceUtf8
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "3.0.5", IconResourceID = 400)]
    [Guid("F0D764E0-A07B-4BAC-B6AE-A1ED025F9525")]
    [ProvideOptionPage(typeof(OptionConfig), "Force UTF8", "General", 0, 0, true)]
    [ProvideAutoLoad("{ADFC4E65-0397-11D1-9F4E-00A0C911004F}", PackageAutoLoadFlags.BackgroundLoad)] // UIContextGuids.EmptySolution
    [ProvideAutoLoad("{ADFC4E64-0397-11D1-9F4E-00A0C911004F}", PackageAutoLoadFlags.BackgroundLoad)] // UIContextGuids.NoSolution
    [ProvideAutoLoad("{F1536EF8-92EC-443C-9ED7-FDADF150DA82}", PackageAutoLoadFlags.BackgroundLoad)] // UIContextGuids.SolutionExists
    public sealed class MainPackage : AsyncPackage
    {        
        public class OptionConfig : DialogPage
        {
            private bool mBom = false;
            private ELineEnding mLineEnding = ELineEnding.Unix;
            private bool mOutput = false;

            [DisplayName("Utf8 Bom")]
            [Description("Contain Bom")]
            public bool Bom
            {
                get { return mBom; }
                set { mBom = value; }
            }

            [DisplayName("Enable Output")]
            [Description("Enable Output")]
            public bool Output
            {
                get { return mOutput; }
                set { mOutput = value; }
            }

            [DisplayName("Line Endings")]
            [Description("Line Endigns")]
            public ELineEnding LineEnding
            {
                get { return mLineEnding; }
                set { mLineEnding = value; }
            }
        }

        public enum ELineEnding
        {
            None,       //
            Unix,       // \n
            Window,     // \r\n
            Mac,        // \r
        }

        public static System.Text.RegularExpressions.Regex RegLineEnding = new System.Text.RegularExpressions.Regex(@"\r\n?|\n");

        private OptionConfig mConfig;
        private DTE mDte;
        private OutputWindowPane mOutputWindowPanel;
        private Events mDteEvents;
        private DocumentEvents mDocumentEvents;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            mConfig = (OptionConfig)GetDialogPage(typeof(OptionConfig));

            mDte = await this.GetServiceAsync(typeof(DTE)) as DTE;
            if (mDte == null)
                return;

            if (mConfig.Output)
            {
                mOutputWindowPanel = _CreateOutputPanel(mDte);
                if (mOutputWindowPanel == null)
                    return;
            }

            //we must save these two references
            // https://social.msdn.microsoft.com/Forums/en-US/0857a868-e650-42ed-b9cc-2975dc46e994/addin-documentevents-are-not-triggered?forum=vsx
            mDteEvents = mDte.Events;
            mDocumentEvents = mDteEvents.DocumentEvents;

            mDocumentEvents.DocumentSaved += DocumentEvents_DocumentSaved;

            _Log("Init Succ \n");
        }

        private void DocumentEvents_DocumentSaved(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _Log("On Document Saved");

            if (document.Kind != "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}")
            {
                _Log($"Document Kind is not correct, {document.Kind}");
                return;
            }

            //Get TextDocument
            TextDocument txt_doc = (TextDocument)document.Object("TextDocument");
            if (txt_doc == null)
            {
                _Log("Failed to get TextDocument");
                return;
            }

            // Get Text Content
            EditPoint editPoint = txt_doc.CreateEditPoint(txt_doc.StartPoint);
            string text = editPoint.GetText(txt_doc.EndPoint);


            //Replace Line Ending
            text = ReplaceLineEndings(text, mConfig.LineEnding);

            //Save
            string path = document.FullName;
            File.WriteAllText(path, text, new UTF8Encoding(mConfig.Bom));

            _Log($"Succ, Bom: {mConfig.Bom}, LineEnding: {mConfig.LineEnding},  Path: {path} \n");
        }

        private void _Log(string msg)
        {
            if (!mConfig.Output)
                return;

            if (mOutputWindowPanel == null)
            {
                mOutputWindowPanel = _CreateOutputPanel(mDte);
                if (mOutputWindowPanel == null)
                    return;
            }

            ThreadHelper.ThrowIfNotOnUIThread();

            mOutputWindowPanel.OutputString(msg);
            mOutputWindowPanel.OutputString("\n");
        }

        private static string ReplaceLineEndings(string text, ELineEnding line_ending)
        {
            switch (line_ending)
            {
                default:
                case ELineEnding.None:
                    //Do Nothing
                    return text;

                case ELineEnding.Window:
                    return RegLineEnding.Replace(text, "\r\n");

                case ELineEnding.Unix:
                    return RegLineEnding.Replace(text, "\n");

                case ELineEnding.Mac:
                    return RegLineEnding.Replace(text, "\r");
            }
        }

        private static OutputWindowPane _CreateOutputPanel(DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            OutputWindow output_window = window.Object as OutputWindow;

            return output_window.OutputWindowPanes.Add("Force Utf8");
        }
    }
}
