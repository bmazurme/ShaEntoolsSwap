using System.Windows;
using MaterialDesignThemes.Wpf;
using MaterialDesignColors;
using System.Windows.Media;


namespace Entools.Plumbing.View
{
    /// <summary>
    /// Логика взаимодействия для UserControl1.xaml
    /// </summary>
    public partial class ErrorWindow : Window
    {
        private void InitializeMaterialDesign()
        {
            // Create dummy objects to force the MaterialDesign assemblies to be loaded
            // from this assembly, which causes the MaterialDesign assemblies to be searched
            // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
            // are searched relative to Eclipse's path, so they're not found.
            var card = new Card();
            var hue = new Hue("Dummy", Colors.Black, Colors.White);
        }

        public ErrorWindow(string Txt, string Title)
        {
            InitializeMaterialDesign();
            InitializeComponent();

            this.Title = Title;
            ErrorContext.Text = Txt;
        }
        public ErrorWindow(string Txt)
        {
            this.Title = Title;
            ErrorContext.Text = Txt;

            InitializeComponent();
            ErrorContext.Text = Txt;
        }
    }
}
