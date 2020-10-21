using System.Windows;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using System.Windows.Media;
using Entools.ViewModels;
using System;
using Entools.Model;

namespace Entools.Views
{
    public partial class WindowScheme : Window
    {
        private void InitializeMaterialDesign()
        {
            // Create dummy objects to force the MaterialDesign assemblies to be loaded
            // from this assembly, which causes the MaterialDesign assemblies to be searched
            // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
            // are searched relative to Eclipse's path, so they're not found.
            var card = new Card();
            var hue = new MaterialDesignColors.Hue("Dummy", Colors.Black, Colors.White);
            var dummy = typeof(MaterialDesignThemes.Wpf.Theme);
        }


        /// <summary>
        /// Show result in window
        /// </summary>
        /// <param name="listSizes"></param>
        /// <param name="listLength"></param>
        public WindowScheme()//List<double> listSizes, List<double> listLength)
        {
            ViewModel viewModel = new ViewModel();
            InitializeMaterialDesign();
            InitializeComponent();
            //Transfer.progressBar = this.progressBar;
            DataContext = viewModel;
            viewModel.CloseAction = new Action(() => this.Close());
        }


        // Close window - ESC
        private void OnCloseExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }
    }
}