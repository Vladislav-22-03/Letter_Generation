using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Letter_Laba_
{
    public partial class MainWindow : Window
    {
        private List<Tuple<TextBox, TextBox>> attachments = new List<Tuple<TextBox, TextBox>>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddAttachment_Click(object sender, RoutedEventArgs e)
        {
            // Контейнер всего приложения
            StackPanel appContainer = new StackPanel
            {
                Margin = new Thickness(0, 5, 0, 5),
                Orientation = Orientation.Vertical
            };

            // Название приложения — текст + мусорка
            TextBlock titleLabel = new TextBlock
            {
                Text = "Название приложения:",
                Margin = new Thickness(0, 5, 0, 0)
            };

            TextBox titleBox = new TextBox
            {
                Height = 25,
                Width = 350,
                Margin = new Thickness(0, 0, 5, 0),
                Tag = "Введите название приложения"
            };

            // Очистка названия
            Button clearTitleBtn = new Button
            {
                Content = "🗑",
                Width = 30,
                Margin = new Thickness(5, 0, 0, 0),
                Tag = titleBox
            };
            clearTitleBtn.Click += (s, ev) => titleBox.Text = "";

            // Горизонтальный контейнер для названия + мусорки
            StackPanel titlePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            titlePanel.Children.Add(titleBox);
            titlePanel.Children.Add(clearTitleBtn);

            // Текст приложения — текст + мусорка
            TextBlock contentLabel = new TextBlock
            {
                Text = "Текст приложения:",
                Margin = new Thickness(0, 5, 0, 0)
            };

            TextBox contentBox = new TextBox
            {
                Height = 60,
                Width = 350,
                Margin = new Thickness(0, 0, 5, 0),
                TextWrapping = TextWrapping.Wrap,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Tag = "Введите текст приложения"
            };

            // Очистка текста приложения
            Button clearContentBtn = new Button
            {
                Content = "🗑",
                Width = 30,
                Margin = new Thickness(5, 0, 0, 0),
                Tag = contentBox
            };
            clearContentBtn.Click += (s, ev) => contentBox.Text = "";

            // Горизонтальный контейнер для текста + мусорки
            StackPanel contentPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            contentPanel.Children.Add(contentBox);
            contentPanel.Children.Add(clearContentBtn);

            // Кнопка удаления всего приложения
            Button removeAttachmentBtn = new Button
            {
                Content = "Удалить приложение",
                Margin = new Thickness(0, 5, 0, 0),
                Background = Brushes.LightCoral
            };
            removeAttachmentBtn.Click += (s, ev) =>
            {
                AttachmentPanel.Children.Remove(appContainer);
                attachments.RemoveAll(t => t.Item1 == titleBox && t.Item2 == contentBox);
            };

            // Собираем всё вместе
            appContainer.Children.Add(titleLabel);
            appContainer.Children.Add(titlePanel);
            appContainer.Children.Add(contentLabel);
            appContainer.Children.Add(contentPanel);
            appContainer.Children.Add(removeAttachmentBtn);

            AttachmentPanel.Children.Add(appContainer);

            // Добавляем в список для вставки в документ
            attachments.Add(new Tuple<TextBox, TextBox>(titleBox, contentBox));
        }





        private void ReplacePlaceholders(string sourceFilePath, string destinationFilePath)
        {
            File.Copy(sourceFilePath, destinationFilePath, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(destinationFilePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                int replacements = 0;
                string greeting = cmbGender.SelectedIndex == 0 ? "Уважаемый " : "Уважаемая ";

                foreach (var paragraph in body.Descendants<Paragraph>().ToList())
                {
                    foreach (var text in paragraph.Descendants<Text>().ToList())
                    {
                        if (text.Text.Contains("[Текст письма]"))
                        {
                            var parentParagraph = text.Ancestors<Paragraph>().FirstOrDefault();
                            if (parentParagraph != null)
                            {
                                parentParagraph.RemoveAllChildren<Run>();

                                // Добавляем текст письма построчно
                                foreach (var line in txtBody.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
                                {
                                    parentParagraph.AppendChild(new Run(new Text(line)));
                                    parentParagraph.AppendChild(new Run(new Break()));
                                }

                                // Добавляем список приложений сразу после текста письма
                                if (attachments.Count > 0)
                                {
                                    parentParagraph.AppendChild(new Run(new Break()));
                                    parentParagraph.AppendChild(new Run(new Text("Приложения:")));
                                    parentParagraph.AppendChild(new Run(new Break()));

                                    for (int i = 0; i < attachments.Count; i++)
                                    {
                                        string title = attachments[i].Item1.Text;
                                        parentParagraph.AppendChild(new Run(new Text($"{i + 1}. {title} (на странице {2 + i})")));
                                        parentParagraph.AppendChild(new Run(new Break()));
                                    }
                                }

                                replacements++;
                            }
                        }
                        else
                        {
                            // Простая замена остальных плейсхолдеров
                            text.Text = text.Text
                                .Replace("[Почта]", txtAddress.Text)
                                .Replace("[Адресат1]", txtRecipient.Text)
                                .Replace("[Должность адресата]", greeting + txtRecipientPost.Text)
                                .Replace("[Адресат]", greeting + txtRecipient.Text)
                                .Replace("[Тема письма]", txtSubject.Text)
                                .Replace("[ФИО]", txtFullName.Text)
                                .Replace("[Должность]", txtPosition.Text)
                                .Replace("[Дата]", DateTime.Now.ToShortDateString());
                        }
                    }
                }

                // Добавляем приложения на отдельные страницы
                if (attachments.Count > 0)
                {
                    for (int i = 0; i < attachments.Count; i++)
                    {
                        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                        body.AppendChild(new Paragraph(
                            new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                            new Run(new RunProperties(new Bold()), new Text($"Приложение {i + 1}"))
                        ));

                        body.AppendChild(new Paragraph(
                            new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                            new Run(new RunProperties(new Bold()), new Text(attachments[i].Item1.Text))
                        ));

                        body.AppendChild(new Paragraph(
                            new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                            new Run(new Text(attachments[i].Item2.Text))
                        ));
                    }
                }

                if (replacements > 0)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else
                {
                    MessageBox.Show("Плейсхолдеры не найдены. Проверьте шаблон.");
                }
            }
        }



        private void ClearTextBox_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            if (button != null && button.Tag is string)
            {
                string textBoxName = button.Tag.ToString();
                TextBox field = this.FindName(textBoxName) as TextBox;
                if (field != null)
                {
                    field.Text = string.Empty;
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Проверка на незаполненные поля
            List<string> emptyFields = new List<string>();

            if (string.IsNullOrWhiteSpace(txtAddress.Text)) emptyFields.Add("Почта");
            if (string.IsNullOrWhiteSpace(txtRecipient.Text)) emptyFields.Add("Адресат");
            if (string.IsNullOrWhiteSpace(txtRecipientPost.Text)) emptyFields.Add("Должность адресата");
            if (string.IsNullOrWhiteSpace(txtSubject.Text)) emptyFields.Add("Тема письма");
            if (string.IsNullOrWhiteSpace(txtBody.Text)) emptyFields.Add("Текст письма");
            if (string.IsNullOrWhiteSpace(txtFullName.Text)) emptyFields.Add("ФИО");
            if (string.IsNullOrWhiteSpace(txtPosition.Text)) emptyFields.Add("Должность");
            if (cmbGender.SelectedIndex == -1) emptyFields.Add("Пол");

            if (emptyFields.Count > 0)
            {
                string message = "Пожалуйста, заполните следующие поля:\n\n" + string.Join("\n", emptyFields);
                MessageBox.Show(message, "Не заполнены поля", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Путь к шаблону
            string sourceFilePath = @"C:\Users\vladi\Desktop\Лаба1.docx";

            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("Файл шаблона не найден!");
                return;
            }

            // Диалог сохранения
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Word Documents|*.docx",
                FileName = "Лаба1_готовый.docx",
                DefaultExt = ".docx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string destinationFilePath = saveFileDialog.FileName;

                try
                {
                    ReplacePlaceholders(sourceFilePath, destinationFilePath);
                    MessageBox.Show("Документ успешно сохранён!");
                }
                catch (IOException)
                {
                    MessageBox.Show("Ошибка: файл открыт в другой программе. Закройте его и попробуйте снова.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при создании документа: " + ex.Message);
                }
            }
        }
    }
}
