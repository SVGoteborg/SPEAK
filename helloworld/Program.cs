using System;
using System.IO;
//using System.Net.Mail;
using System.Threading.Tasks;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;



class Program
{
    private const string FileFromPath = "C:\\Users\\admin\\Desktop\\Praktik\\Speechtotext";
    private const string FileFromName = "MSG00001.wav";
    private const string FileToRead = FileFromPath + "\\" + FileFromName;
    private const string FileToPath = "C:\\Users\\admin\\Desktop\\Result";
    private const string FileToName = "WriteText2.txt";
    private const string FileToWrite = FileToPath + "\\" + FileToName;
    private static string TranscribedMessage = null;


    async static Task FromFile(SpeechConfig speechConfig)
    {
        using var audioConfig = AudioConfig.FromWavFileInput(FileToRead);
        //using var audioConfig = AudioConfig.FromWavFileInput("PathToFile.wav");
        using var recognizer = new SpeechRecognizer(speechConfig, audioConfig);

        var result = await recognizer.RecognizeOnceAsync();
        Console.WriteLine($"RECOGNIZED: Text={result.Text}");
        SaveText(result.Text);
        TranscribedMessage = result.Text;
    }

    static void SaveText(string text)
    {
        File.WriteAllText(FileToWrite, text);
    }

    async static Task Main(string[] args)
    {
        var speechConfig = SpeechConfig.FromSubscription("ad600bd7bbb84bda813532091b74fd2b", "westeurope");
        speechConfig.SpeechRecognitionLanguage = "sv-SE";
        await FromFile(speechConfig);

        // Send email
        var message = new MimeMessage();
        message.From.Add(new MailboxAddress("Avvikelse", "avvikelse@testboka.net"));
        message.To.Add(new MailboxAddress("Avvikelse2", "avvikelse2@testboka.net"));
        message.Subject = "Transkriberad text";

        message.Body = new TextPart("plain")
        {
            Text = TranscribedMessage
        };

        using (var client = new SmtpClient())
        {
            client.Connect("smtp.simply.com", 587, false);

            // Note: only needed if the SMTP server requires authentication
            client.Authenticate("avvikelse@testboka.net", "Elektronik!100");

            client.Send(message);
            client.Disconnect(true);
        }
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

