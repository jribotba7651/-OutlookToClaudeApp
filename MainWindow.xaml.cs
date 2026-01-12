using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Threading.Tasks;
using OutlookToClaudeApp.Models;
using OutlookToClaudeApp.Services;

namespace OutlookToClaudeApp
{
    public partial class MainWindow : Window
    {
        private List<CalendarEvent> _allEvents = new List<CalendarEvent>();
        private OutlookServiceV3 _outlookService;
        private ClaudeApiService _claudeService;

        public MainWindow()
        {
            InitializeComponent();
            InitializeDefaults();
        }

        private void InitializeDefaults()
        {
            // Set default date range to current month
            StartDatePicker.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            EndDatePicker.SelectedDate = StartDatePicker.SelectedDate.Value.AddMonths(1).AddDays(-1);
        }

        private async void LoadEventsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!StartDatePicker.SelectedDate.HasValue || !EndDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Please select both start and end dates.", "Invalid Date Range",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                LoadEventsButton.IsEnabled = false;
                StatusText.Text = "Loading events from Outlook...";

                _outlookService?.Dispose();
                _outlookService = new OutlookServiceV3();

                _allEvents = _outlookService.GetEvents(
                    StartDatePicker.SelectedDate.Value,
                    EndDatePicker.SelectedDate.Value
                );

                EventsListBox.ItemsSource = null;
                EventsListBox.ItemsSource = _allEvents;

                StatusText.Text = $"Loaded {_allEvents.Count} events";

                if (_allEvents.Count == 0)
                {
                    MessageBox.Show("No events found in the selected date range.",
                        "No Events", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load events: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Error loading events";
            }
            finally
            {
                LoadEventsButton.IsEnabled = true;
            }
        }


        private void EventsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EventsListBox.SelectedItem is CalendarEvent selectedEvent)
            {
                DisplayEventDetails(selectedEvent);
            }
        }

        private void DisplayEventDetails(CalendarEvent evt)
        {
            DetailsPanel.Children.Clear();

            // Subject
            var subjectBlock = new TextBlock
            {
                Text = evt.DisplayTitle,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 15),
                TextWrapping = TextWrapping.Wrap
            };
            DetailsPanel.Children.Add(subjectBlock);

            // Date
            AddDetailRow("Date:", evt.DisplayDate);

            // Time
            AddDetailRow("Time:", evt.DisplayTime);

            // Location
            if (!string.IsNullOrWhiteSpace(evt.Location))
            {
                AddDetailRow("Location:", evt.Location);
            }

            // Organizer
            if (!string.IsNullOrWhiteSpace(evt.Organizer))
            {
                AddDetailRow("Organizer:", evt.Organizer);
            }

            // Categories
            if (!string.IsNullOrWhiteSpace(evt.Categories))
            {
                AddDetailRow("Categories:", evt.Categories);
            }

            // Body
            if (!string.IsNullOrWhiteSpace(evt.Body))
            {
                var separator = new Border
                {
                    Height = 1,
                    Background = new SolidColorBrush(Color.FromRgb(224, 224, 224)),
                    Margin = new Thickness(0, 15, 0, 15)
                };
                DetailsPanel.Children.Add(separator);

                var bodyLabel = new TextBlock
                {
                    Text = "Details:",
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 5)
                };
                DetailsPanel.Children.Add(bodyLabel);

                var bodyBlock = new TextBlock
                {
                    Text = evt.Body.Trim(),
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102))
                };
                DetailsPanel.Children.Add(bodyBlock);
            }
        }

        private void AddDetailRow(string label, string value)
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8)
            };

            var labelBlock = new TextBlock
            {
                Text = label,
                FontWeight = FontWeights.SemiBold,
                Width = 100
            };

            var valueBlock = new TextBlock
            {
                Text = value,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102))
            };

            panel.Children.Add(labelBlock);
            panel.Children.Add(valueBlock);
            DetailsPanel.Children.Add(panel);
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var evt in _allEvents)
            {
                evt.IsSelected = true;
            }
            EventsListBox.Items.Refresh();
            UpdateStatusCount();
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var evt in _allEvents)
            {
                evt.IsSelected = false;
            }
            EventsListBox.Items.Refresh();
            UpdateStatusCount();
        }

        private void UpdateStatusCount()
        {
            var selectedCount = _allEvents.Count(e => e.IsSelected);
            StatusText.Text = $"Selected {selectedCount} of {_allEvents.Count} events";
        }

        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedEvents = _allEvents.Where(e => e.IsSelected).ToList();

            if (selectedEvents.Count == 0)
            {
                MessageBox.Show("Please select at least one event to preview.",
                    "No Events Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var tempService = new ClaudeApiService("temp-key");
            var content = tempService.GenerateMarkdown(selectedEvents);

            var previewWindow = new Window
            {
                Title = "Markdown Preview",
                Width = 800,
                Height = 600,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };

            var textBox = new TextBox
            {
                Text = content,
                IsReadOnly = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"),
                Padding = new Thickness(10)
            };

            previewWindow.Content = textBox;
            previewWindow.ShowDialog();
        }

        private async void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedEvents = _allEvents.Where(ev => ev.IsSelected).ToList();

            if (selectedEvents.Count == 0)
            {
                MessageBox.Show("Please select at least one event to export.",
                    "No Events Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var apiKey = ApiKeyBox.Password;
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show("Please enter your Claude API Key.",
                    "Missing API Key", MessageBoxButton.OK, MessageBoxImage.Warning);
                ApiKeyBox.Focus();
                return;
            }

            try
            {
                ExportButton.IsEnabled = false;
                StatusText.Text = $"Uploading {selectedEvents.Count} events to Claude...";

                // Initialize service with real key
                _claudeService = new ClaudeApiService(apiKey);

                // Upload
                var result = await _claudeService.UploadCalendarAsync(selectedEvents);

                if (result.Success)
                {
                    // Copy to clipboard
                    Clipboard.SetText(result.FileId);

                    var message = $"Successfully uploaded to Claude!\n\n" +
                                  $"File ID: {result.FileId}\n\n" +
                                  "The File ID has been copied to your clipboard.\n" +
                                  "You can now paste it into your conversation with Claude.";

                    MessageBox.Show(message, "Upload Successful",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    StatusText.Text = "Upload successful! File ID copied to clipboard.";
                }
                else
                {
                    MessageBox.Show(result.Message, "Upload Failed",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "Upload failed.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to upload: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Error uploading events";
            }
            finally
            {
                ExportButton.IsEnabled = true;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _outlookService?.Dispose();
            base.OnClosed(e);
        }
    }
}