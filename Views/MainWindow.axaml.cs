using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;

namespace MethodExpertSurveys.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void ChooseExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var topLevel = GetTopLevel(this);
            var files = await topLevel!.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Open Excel File",
                AllowMultiple = false,
                FileTypeFilter = [new("Excel Files") { Patterns = ["*.xlsx", "*.xls"] } ]
            });
  
            if (files.Count > 0)
                ExcelFilePath.Text = files[0].Path.LocalPath;
        }

        private string format = string.Empty;
        private void MethodList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox { SelectedItem: ListBoxItem selectedItem } && selectedItem.Content is string content)
            {
                format = content;
                selectedItem.FontWeight = Avalonia.Media.FontWeight.Bold;
            }
        }

        private async void ChooseOutputFilePath_Click(object sender, RoutedEventArgs e) 
        {
            var topLevel = GetTopLevel(this);
            var folder = await topLevel!.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Open Path"
            });

            if (folder.Count > 0)
                OutputFilePath.Text = folder[0].Path.LocalPath;
        }

        private void Execute_Click(object sender, RoutedEventArgs e) 
        {
            var data = (ExcelFilePath.Text!, TableRange.Text!, OutputFilePath.Text!);
            switch (format)
            {
                case "Direct Ranking":
                    RankingBuilder.DirectRankingBuild(data.Item1, data.Item2, data.Item3);
                    WarningText.Text = "Файл по Direct Ranking сгенерирован!";
                    break;
                case "Pairwise Comparison Ranking":
                    RankingBuilder.PairComRankingBuild(data.Item1, data.Item2, data.Item3);
                    WarningText.Text = "Файл по Pairwise Comparison Ranking сгенерирован!";
                    break;
            }
        }

    }
}