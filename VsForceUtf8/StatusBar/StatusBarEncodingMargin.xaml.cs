using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text;
using System.Windows.Input;
using System.Text;

namespace VsForceUtf8
{
    /// <summary>
    /// Margin's canvas and visual definition including both size and content
    /// </summary>
    public partial class StatusBarEncodingMargin : Button, IWpfTextViewMargin
    {
        private static Encoding Encoding_Unicode = Encoding.Unicode;
        private static Encoding Encoding_BEUnicode = Encoding.BigEndianUnicode;
        private static Encoding Encoding_Utf8 = Encoding.UTF8;
        private static Encoding Encoding_Utf8NoBom = new UTF8Encoding(false);
        private static Encoding Encoding_Default = Encoding.Default;

        //GB18030 > GBK > GB2312
        private static Encoding Encoding_GB2312 = Encoding.GetEncoding("GB18030");
        private static Encoding Encoding_GBK = Encoding.GetEncoding("GBK");
        private static Encoding Encoding_GB18030 = Encoding.GetEncoding("GB2312");


        private struct InnerMenuItem
        {
            public string Name;
            public Encoding Encoding;
            public InnerMenuItem(Encoding encoding, string name)
            {
                this.Name = name;
                this.Encoding = encoding;
            }
        }

        private static InnerMenuItem[] _ConvertMenus = new InnerMenuItem[]
        {
            new InnerMenuItem(Encoding_Utf8NoBom, "UTF-8"),
            new InnerMenuItem(Encoding_Utf8,"UTF-8 BOM"),
            new InnerMenuItem(Encoding_GB18030, "GB18030"),
        };

        /// <summary>
        /// Margin name.
        /// </summary>
        public const string MarginName = "StatusBarEncodeMargin";

        /// <summary>
        /// A value indicating whether the object is disposed.
        /// </summary>
        private bool _isDisposed;

        /// <summary>
        /// The document of textview.
        /// </summary>
        private readonly ITextDocument _document = null;

        internal class ConvertCommand : ICommand
        {
            private readonly Encoding _encoding;
            public ConvertCommand(Encoding encoding)
            {
                this._encoding = encoding;
            }

            event EventHandler ICommand.CanExecuteChanged
            {
                add { }
                remove { }
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public void Execute(object parameter)
            {
                StatusBarEncodingMargin self = parameter as StatusBarEncodingMargin;
                if (_IsEncodingSame(self._document.Encoding, this._encoding))
                    return;

                self._document.Encoding = _encoding;
                self._document.UpdateDirtyState(true, DateTime.Now);
                self.Content = _GetDocumentEncoding(self._document);
            }
        }


        /// <summary>
        /// Initializes a new instance.
        /// </summary>
        /// <param name="textView">The textView to attach the margin to.</param>
        public StatusBarEncodingMargin(IWpfTextView textView, IWpfTextViewMargin marginContainer)
        {
            InitializeComponent();
            // display
            ClipToBounds = true;

            // Text
            if (!textView.TextBuffer.Properties.TryGetProperty(typeof(ITextDocument), out _document))
                textView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument), out _document);


            Content = _GetDocumentEncoding(_document);

            // Menu
            ContextMenu = new ContextMenu();


            foreach (var menu in _ConvertMenus)
            {
                //string text = "Convert to " + menu.Name;
                _ = ContextMenu.Items.Add(new MenuItem
                {
                    Header = "Convert To " + menu.Name,
                    Command = new ConvertCommand(menu.Encoding),
                    CommandParameter = this
                });
            } 

            ContextMenu.PlacementTarget = this;
            ContextMenu.Placement = PlacementMode.Top;
            Click += _OnStatusBar_Click;
            _document.FileActionOccurred += (sender, e) => Content = _GetDocumentEncoding(_document);
        }

        private void _OnStatusBar_Click(object sender, RoutedEventArgs e)
        {
            var encodingNow = this._document.Encoding;
            for (int i = 0; i < ContextMenu.Items.Count; ++i)
            {
                MenuItem item = ContextMenu.Items[i] as MenuItem;
                item.IsChecked = _IsEncodingSame(encodingNow, _ConvertMenus[i].Encoding);
            }             

            ContextMenu.IsOpen = true;
        }

        private static bool _IsEncodingSame(Encoding e1, Encoding e2)
        {
            if (e1 == null && e2 == null)
                return true;
            else if (e1 == null || e2 == null)
                return false;

            if (e1.CodePage != e2.CodePage)
                return false;
            return e1.GetPreamble().Length == e2.GetPreamble().Length;
        }

        private static string _GetDocumentEncoding(ITextDocument document)
        {
            if (document == null || document.Encoding == null)
                return "Unkown";

            var encodingNow = document.Encoding;
            foreach (var p in _ConvertMenus)
            {
                if (_IsEncodingSame(encodingNow, p.Encoding))
                    return p.Name;
            }
            return encodingNow.EncodingName;
        }

        #region AutoGenerate
        #region IWpfTextViewMargin

        /// <summary>
        /// Gets the <see cref="Sytem.Windows.FrameworkElement"/> that implements the visual representation of the margin.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The margin is disposed.</exception>
        public FrameworkElement VisualElement
        {
            // Since this margin implements Canvas, this is the object which renders
            // the margin.
            get
            {
                ThrowIfDisposed();
                return this;
            }
        }

        #endregion

        #region ITextViewMargin

        /// <summary>
        /// Gets the size of the margin.
        /// </summary>
        /// <remarks>
        /// For a horizontal margin this is the height of the margin,
        /// since the width will be determined by the <see cref="ITextView"/>.
        /// For a vertical margin this is the width of the margin,
        /// since the height will be determined by the <see cref="ITextView"/>.
        /// </remarks>
        /// <exception cref="ObjectDisposedException">The margin is disposed.</exception>
        public double MarginSize
        {
            get
            {
                ThrowIfDisposed();

                // Since this is a horizontal margin, its width will be bound to the width of the text view.
                // Therefore, its size is its height.
                return ActualHeight;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the margin is enabled.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The margin is disposed.</exception>
        public bool Enabled
        {
            get
            {
                ThrowIfDisposed();

                // The margin should always be enabled
                return true;
            }
        }

        /// <summary>
        /// Gets the <see cref="ITextViewMargin"/> with the given <paramref name="marginName"/> or null if no match is found
        /// </summary>
        /// <param name="marginName">The name of the <see cref="ITextViewMargin"/></param>
        /// <returns>The <see cref="ITextViewMargin"/> named <paramref name="marginName"/>, or null if no match is found.</returns>
        /// <remarks>
        /// A margin returns itself if it is passed its own name. If the name does not match and it is a container margin, it
        /// forwards the call to its children. Margin name comparisons are case-insensitive.
        /// </remarks>
        /// <exception cref="ArgumentNullException"><paramref name="marginName"/> is null.</exception>
        public ITextViewMargin GetTextViewMargin(string marginName)
        {
            return string.Equals(marginName, MarginName, StringComparison.OrdinalIgnoreCase) ? this : null;
        }

        /// <summary>
        /// Disposes an instance of <see cref="FileEncodeMargin"/> class.
        /// </summary>
        public void Dispose()
        {
            if (!_isDisposed)
            {
                GC.SuppressFinalize(this);
                _isDisposed = true;
            }
        }

        #endregion

        /// <summary>
        /// Checks and throws <see cref="ObjectDisposedException"/> if the object is disposed.
        /// </summary>
        private void ThrowIfDisposed()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(MarginName);
            }
        }
        #endregion
    }
}

