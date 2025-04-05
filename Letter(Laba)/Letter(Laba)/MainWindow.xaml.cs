using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Letter_Laba_
{
    public partial class MainWindow : Window
    {
        private List<TextBox> attachmentTextBoxes = new List<TextBox>(); // Список полей для приложений

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddAttachment_Click(object sender, RoutedEventArgs e)
        {
            // Создаем новое поле для приложения
            TextBox newAttachment = new TextBox
            {
                Height = 60,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 5, 0, 0)
            };

            // Добавляем в список и UI
            attachmentTextBoxes.Add(newAttachment);
            AttachmentPanel.Children.Add(newAttachment);
        }

        private void ReplacePlaceholders(string sourceFilePath, string destinationFilePath)
{
    File.Copy(sourceFilePath, destinationFilePath, true);

    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(destinationFilePath, true))
    {
        var body = wordDoc.MainDocumentPart.Document.Body;
        int replacements = 0;

        string greeting = cmbGender.SelectedIndex == 0 ? "Уважаемый " : "Уважаемая ";

                // Заменяем обычные плейсхолдеры
                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    var textElement = paragraph.Descendants<Text>().FirstOrDefault(t => t.Text.Contains("[Приложение]"));
                    if (textElement != null)
                    {
                        paragraph.RemoveAllChildren<Run>(); // Удаляем старый текст

                        // Создаем новый параграф для списка приложений
                        Paragraph applicationListParagraph = new Paragraph(
                            new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                            new Run(new RunProperties(new Bold()), new Text("Приложения:")),
                            new Run(new Break()) // Разрыв строки после "Приложения:"
                        );

                        for (int i = 0; i < attachmentTextBoxes.Count; i++)
                        {
                            applicationListParagraph.Append(
                                new Run(new Text($"{i + 1}. Приложение {i + 1} (на странице {2 + i})")),
                                new Run(new Break()) // Разрыв строки после каждой записи
                            );
                        }

                        // Вставляем **НОВЫЙ** экземпляр списка
                        paragraph.Append(applicationListParagraph.Elements<Run>().Select(run => new Run(run.OuterXml)));

                        replacements++;
                        break;
                    }
                }


                // Проверяем, есть ли приложения
                if (attachmentTextBoxes.Count > 0)
        {
            int startPage = 2; // Первая страница с приложениями начинается со 2-й

            // Создаем параграф для списка приложений
            Paragraph applicationListParagraph = new Paragraph(
                new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                new Run(new RunProperties(new Bold()), new Text("Приложения:")),
                new Run(new Break()) // Разрыв строки после "Приложения:"
            );

            for (int i = 0; i < attachmentTextBoxes.Count; i++)
            {
                applicationListParagraph.Append(
                    new Run(new Text($"{i + 1}. Приложение {i + 1} (на странице {startPage})")),
                    new Run(new Break()) // Разрыв строки после каждой записи
                );
                startPage++;
            }

            // Заменяем [Приложение] на список
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var textElement = paragraph.Descendants<Text>().FirstOrDefault(t => t.Text.Contains("[Приложение]"));
                if (textElement != null)
                {
                    paragraph.RemoveAllChildren<Run>(); // Очищаем старый текст
                    paragraph.Append(applicationListParagraph.Elements<Run>());
                    replacements++;
                    break;
                }
            }

            // Добавляем каждое приложение на новую страницу
            for (int i = 0; i < attachmentTextBoxes.Count; i++)
            {
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                // Заголовок "Приложение X" (справа)
                body.AppendChild(new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                    new Run(new RunProperties(new Bold()), new Text($"Приложение {i + 1}"))
                ));

                // Заголовок "ПРИЛОЖЕНИЕ" (по центру)
                body.AppendChild(new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                    new Run(new RunProperties(new Bold()), new Text("ПРИЛОЖЕНИЕ"))
                ));

                // Текст приложения (слева)
                body.AppendChild(new Paragraph(
                    new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                    new Run(new Text(attachmentTextBoxes[i].Text))
                ));
            }
        }

        if (replacements > 0)
        {
            wordDoc.MainDocumentPart.Document.Save();
        }
        else
        {
            MessageBox.Show("Ни один плейсхолдер не найден. Проверьте шаблон.");
        }
    }
}

        private void ClearField_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string textBoxName)
            {
                var field = this.FindName(textBoxName) as TextBox;
                if (field != null)
                {
                    field.Text = string.Empty;
                }
            }
        }

        private void ClearText_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string targetName)
            {
                var textBox = FindName(targetName) as TextBox;
                if (textBox != null)
                {
                    textBox.Clear();
                }
            }
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string sourceFilePath = @"C:\Users\vladi\Desktop\Лаба1.docx";

            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("Файл шаблона не найден!");
                return;
            }

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
                    MessageBox.Show($"Ошибка при обновлении документа: {ex.Message}");
                }
            }
        }
    }
}
