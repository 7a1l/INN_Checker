using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Mime;
using System.Linq;
using System.Text.RegularExpressions;
using Avalonia.Controls;
using Avalonia.Interactivity;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Wordprocessing;

namespace INN_Checker;

public partial class MainWindow : Window
{
    private string DataFromAPI = "";
    public MainWindow()
    {
        InitializeComponent();
    }
    private async void GetDataFromAPIBtn_OnClick(object? sender, RoutedEventArgs e)
    {
        var client = new HttpClient();
        var result = await client.GetStringAsync("http://127.0.0.1:4444/TransferSimulator/inn");
        var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(result);
        DataFromAPI = data["value"];
        DataFromAPITextBlock.Text = DataFromAPI;
    }
    private void WriteInDocBtn_OnClick(object? sender, RoutedEventArgs e)
    {
        string regex = @"^\d{10}$";
        var validationResult = Regex.IsMatch(DataFromAPI, regex);
        TestResultTextBlock.Text = validationResult ? "ИНН не содержит запрещенные символы" : "ИНН содержит запрещенные символы";
        
        using var doc = WordprocessingDocument.Open(@"ТестКейс.docx", true);
        var document = doc.MainDocumentPart!.Document;
        
        if (document.Descendants<Text>().Any(text => text.Text.Contains("Result 1")))
        {
            ReplaceText("Result 1", validationResult, document);
        } 
        else if (document.Descendants<Text>().Any(text => text.Text.Contains("Result 2")))
        {
            ReplaceText("Result 2", validationResult, document);
        }
    }
    private void ReplaceText(string replacedText, bool validationResult, Document document)
    {
        foreach (var text in document.Descendants<Text>())
        {
            if (text.Text.Contains(replacedText))
            {
                text.Text = text.Text.Replace(replacedText, validationResult ? "Успешно" : "Не успешно");
            }
        }
    }
    
}