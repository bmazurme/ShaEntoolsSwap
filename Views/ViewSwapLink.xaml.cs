using System.Windows;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using System.Windows.Media;
using System;
using Entools.ViewLinkModels;

namespace Entools.Views
{
    public partial class WindowLink : Window
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
        public WindowLink()
        {
            ViewModelLink viewModel = new ViewModelLink();
            InitializeMaterialDesign();
            InitializeComponent();
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