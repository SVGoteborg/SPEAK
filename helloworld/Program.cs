using System;
using System.IO;
//using System.Net.Mail;
using System.Threading.Tasks;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using MailKit.Net.Pop3;

class Program
{
    //private const string jsonSettings = "C:\\Users\\admin\\Desktop\\AppSpeak\\settings.json";
    private const string jsonSettings = "C:\\AppSpeak\\settings.json";
    private static string FileFromPath = "C:\\AppSpeak\\ljud";
    //private static string FileFromPath = "C:\\Users\\admin\\Desktop\\AppSpeak\\ljud";
    private static string FileFromName;

    //Kommer inte att behövas -----
    private static string FileFromName2 = "MSG00001.wav";
    private static string FileToSave = FileFromPath + "\\" + FileFromName2;
    private const string FileToPath = "C:\\Users\\admin\\Desktop\\AppSpeak\\Result";
    private const string FileToName = "WriteText2.txt";
    private const string FileToWrite = FileToPath + "\\" + FileToName;
    // --------

    private static MimeMessage eMail = null;
    private static string TranscribedMessage = null;
    private static string fromName = "Avvikelse";
    private static string fromEmail = "avvikelse@testboka.net";
    private static string fromClient = "mail.simply.com";
    private static int fromPort = 110;
    private static int fromImapPort = 143;
    private static string toClient = "smtp.simply.com";
    private static int toPort = 587;
    private static string passw = "Elektronik!100";
    private static string toName = "Avvikelse2";
    private static string toEmail = "avvikelse2@testboka.net";
    private static string AzureSubscription = "ad600bd7bbb84bda813532091b74fd2b";
    private static string AzureServer = "westeurope";

    public static MailboxAddress Sender { get; private set; }
    public static MailboxAddress Recipient { get; private set; }

    public static MimeMessage Forward(MimeMessage original, MailboxAddress from, InternetAddress to)
    {
        var message = new MimeMessage();
        message.From.Add(from);
        message.To.Add(to);

        // set the forwarded subject
        if (!original.Subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase))
            message.Subject = "FW: " + original.Subject;
        else
            message.Subject = original.Subject;

        // create the main textual body of the message
        var text = new TextPart("plain") { Text = "Here's the forwarded message: Test för forward av mail" };

        // create the message/rfc822 attachment for the original message
        var rfc822 = new MessagePart { Message = original };

        // create a multipart/mixed container for the text body and the forwarded message
        var multipart = new Multipart("mixed");
        multipart.Add(text);
        multipart.Add(rfc822);

        // set the multipart as the body of the message
        message.Body = multipart;

        return message;
    }

    async static Task FromFile(SpeechConfig speechConfig)
    {
        //var fileToSave = FileToSave;
        string fileToSave;

        using (var client = new Pop3Client())
        {
            client.Connect(fromClient, fromPort, false);

            client.Authenticate(fromEmail, passw);

            for (int i = 0; i < client.Count; i++)
            {
                eMail = client.GetMessage(i);

                foreach (var attachment in eMail.Attachments)
                {
                    if (attachment is MessagePart)
                    {
                        var fileName = attachment.ContentDisposition?.FileName;
                        var rfc822 = (MessagePart)attachment;

                        if (string.IsNullOrEmpty(fileName))
                            fileName = "attached-message.eml";

                        using (var stream = File.Create(fileName))
                            rfc822.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)attachment;
                        var fileName = part.FileName;

                        using (var stream = File.Create(fileName))
                        {
                            part.Content.DecodeTo(stream);
                        }
                        // Här ändra till rätt filnamn
                        fileToSave = FileFromPath + "\\" + fileName;
                        using (var stream = File.Create(fileToSave))
                        {
                            part.Content.DecodeTo(stream);
                        }
                    }
                }
                Console.WriteLine("Subject: {0}", eMail.Subject);
                // Forwarding
                Sender = new MailboxAddress(fromName, fromEmail);
                Recipient = new MailboxAddress(toName, toEmail);
                var mailToForward = Forward(eMail, Sender, Recipient);
                using (var smtp = new MailKit.Net.Smtp.SmtpClient())
                {
                    await smtp.ConnectAsync(toClient, toPort, false);
                    await smtp.AuthenticateAsync(fromEmail, passw);
                    await smtp.SendAsync(mailToForward);
                    await smtp.DisconnectAsync(true);
                }
            }
            client.Disconnect(true);
        }

        //string[] fileArray = Directory.GetFiles(FileFromPath, "*.wav");
        //FileFromName = fileArray[0];
        //using var audioConfig = AudioConfig.FromWavFileInput(FileFromName);
        ////using var audioConfig = AudioConfig.FromWavFileInput("PathToFile.wav");
        //using var recognizer = new SpeechRecognizer(speechConfig, audioConfig);

        //var result = await recognizer.RecognizeOnceAsync();
        //Console.WriteLine($"RECOGNIZED: Text={result.Text}");
        //// Skall tas bort
        //SaveText(result.Text);

        //TranscribedMessage = result.Text;

    }
    // Skall tas bort
    //static void SaveText(string text)
    //{
    //    File.WriteAllText(FileToWrite, text);
    //}

    static void checkFolder()
    {

    }

    static void checkJson()
    {

    }

    static void readJson()
    {
        JObject o1 = JObject.Parse(File.ReadAllText(jsonSettings));

        // För test
        foreach (JProperty property in o1.Properties())
        {
            Console.WriteLine(property.Name + " - " + property.Value);
        }
        Console.WriteLine();
        // -------

        FileFromPath = (string)o1["FileFromPath"];
        fromName = (string)o1["fromName"];
        fromEmail = (string)o1["fromEmail"];
        fromClient = (string)o1["fromClient"];
        fromPort = (int)o1["fromPort"];
        fromImapPort = (int)o1["fromImapPort"];
        toClient = (string)o1["toClient"];
        toPort = (int)o1["toPort"];
        passw = (string)o1["passw"];
        toName = (string)o1["toName"];
        AzureSubscription = (string)o1["AzureSubscription"];
        AzureServer = (string)o1["AzureServer"];

    }
    

async static Task Main(string[] args)
    {
        checkFolder();
        checkJson();
        readJson();
        var speechConfig = SpeechConfig.FromSubscription(AzureSubscription, AzureServer);
        speechConfig.SpeechRecognitionLanguage = "sv-SE";
        await FromFile(speechConfig);

        // Send email
        //var message = new MimeMessage();
        //message.From.Add(new MailboxAddress(fromName, fromEmail));
        //message.To.Add(new MailboxAddress(toName, toEmail));
        //message.Subject = "Transkriberad text";

        //message.Body = new TextPart("plain")
        //{
        //    Text = TranscribedMessage
        //};

        //using (var client = new SmtpClient())
        //{
        //    client.Connect(toClient, toPort, false);

        //    // Note: only needed if the SMTP server requires authentication
        //    client.Authenticate(fromEmail, passw);

        //    client.Send(message);
        //    client.Disconnect(true);
        //}
        //DirectoryInfo di = new DirectoryInfo(FileFromPath);

        //foreach (FileInfo file in di.GetFiles())
        //{
        //    file.Delete();
        //}
    }

    
}

//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE.md file in the project root for full license information.
//

// <code>
//using System;
//using System.Threading.Tasks;
//using Microsoft.CognitiveServices.Speech;
//using Microsoft.CognitiveServices.Speech.Translation;

//namespace helloworld
//{
//    class Program
//    {
//        public static async Task TranslationContinuousRecognitionAsync()
//        {
//            // Creates an instance of a speech translation config with specified subscription key and service region.
//            // Replace with your own subscription key and service region (e.g., "westus").
//            var config = SpeechTranslationConfig.FromSubscription("ee4aaaaa-52a7-4d9e-8e7b-bf60da5213d2", "westeurope");

//            // Sets source and target languages.
//            string fromLanguage = "sv-SE";
//            //string fromLanguage = "en-US";
//            config.SpeechRecognitionLanguage = fromLanguage;
//            config.AddTargetLanguage("sv");
//            //config.AddTargetLanguage("de");

//            // Sets voice name of synthesis output.
//            const string GermanVoice = "sv-SE-SofieNeural";
//            //const string GermanVoice = "de-DE-Hedda";
//            config.VoiceName = GermanVoice;
//            // Creates a translation recognizer using microphone as audio input.
//            using (var recognizer = new TranslationRecognizer(config))
//            {
//                // Subscribes to events.
//                recognizer.Recognizing += (s, e) =>
//                {
//                    Console.WriteLine($"RECOGNIZING in '{fromLanguage}': Text={e.Result.Text}");
//                    foreach (var element in e.Result.Translations)
//                    {
//                        Console.WriteLine($"    TRANSLATING into '{element.Key}': {element.Value}");
//                    }
//                };

//                recognizer.Recognized += (s, e) =>
//                {
//                    if (e.Result.Reason == ResultReason.TranslatedSpeech)
//                    {
//                        Console.WriteLine($"\nFinal result: Reason: {e.Result.Reason.ToString()}, recognized text in {fromLanguage}: {e.Result.Text}.");
//                        foreach (var element in e.Result.Translations)
//                        {
//                            Console.WriteLine($"    TRANSLATING into '{element.Key}': {element.Value}");
//                        }
//                    }
//                };

//                recognizer.Synthesizing += (s, e) =>
//                {
//                    var audio = e.Result.GetAudio();
//                    Console.WriteLine(audio.Length != 0
//                        ? $"AudioSize: {audio.Length}"
//                        : $"AudioSize: {audio.Length} (end of synthesis data)");
//                };

//                recognizer.Canceled += (s, e) =>
//                {
//                    Console.WriteLine($"\nRecognition canceled. Reason: {e.Reason}; ErrorDetails: {e.ErrorDetails}");
//                };

//                recognizer.SessionStarted += (s, e) =>
//                {
//                    Console.WriteLine("\nSession started event.");
//                };

//                recognizer.SessionStopped += (s, e) =>
//                {
//                    Console.WriteLine("\nSession stopped event.");
//                };

//                // Starts continuous recognition. Uses StopContinuousRecognitionAsync() to stop recognition.
//                Console.WriteLine("Say something...");
//                await recognizer.StartContinuousRecognitionAsync();

//                do
//                {
//                    Console.WriteLine("Press Enter to stop");
//                } while (Console.ReadKey().Key != ConsoleKey.Enter);

//                // Stops continuous recognition.
//                await recognizer.StopContinuousRecognitionAsync();
//            }
//        }

//        static async Task Main(string[] args)
//        {
//            await TranslationContinuousRecognitionAsync();
//        }
//    }
//}
// </code>

