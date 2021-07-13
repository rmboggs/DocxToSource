using ICSharpCode.AvalonEdit.Search;
using System.Windows;

namespace DocxToSource
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            SearchPanel.Install(this.xXmlSourceEditor);
            SearchPanel.Install(this.xCodeSourceEditor);
        }
    }
}
