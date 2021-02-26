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
using MailKit.Net.Imap;
using MailKit.Search;

class Program
{
    private const string jsonSettings = "C:\\AppSpeak\\settings.json";
    private static string FileFromPath = "C:\\AppSpeak\\ljud";
    private static string FileFromName;
    private static string moveTo = "Transkriberat";

    //Kommer inte att behövas -----
    //private static string FileFromName2 = "MSG00001.wav";
    //private static string FileToSave = FileFromPath + "\\" + FileFromName2;
    //private const string FileToPath = "C:\\Users\\admin\\Desktop\\AppSpeak\\Result";
    //private const string FileToName = "WriteText2.txt";
    //private const string FileToWrite = FileToPath + "\\" + FileToName;
    // --------

    private static MimeMessage eMail = null;
    private static string transcribedMessage = null;
    private static string tempMessage = null;
    private static string fromName = "Avvikelse";
    private static string fromEmail = "avvikelse@testboka.net";
    private static string fromClient = "mail.simply.com";
    private static int fromPort = 110;
    private static int fromImapPort = 143;
    private static string toClient = "smtp.simply.com";
    private static int toPort = 587;
    private static string passw = "XXX";
    private static string toName = "Avvikelse2";
    private static string toEmail = "avvikelse2@testboka.net";
    private static string AzureSubscription = "XXX";
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
        var text = new TextPart("plain") { Text = transcribedMessage };

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

    //async static Task FromFile(SpeechConfig speechConfig)
    async static Task FromFile()
    {
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
                        fileToSave = FileFromPath + "\\" + fileName;
                        using (var stream = File.Create(fileToSave))
                        {
                            part.Content.DecodeTo(stream);
                        }
                    }
                }

                // Transcribing voice message
                //string[] fileArray = Directory.GetFiles(FileFromPath, "*.wav");
                //FileFromName = fileArray[0];
                //using var audioConfig = AudioConfig.FromWavFileInput(FileFromName);
                //using var audioConfig = AudioConfig.FromWavFileInput("PathToFile.wav");
                //using var recognizer = new SpeechRecognizer(speechConfig, audioConfig);

                //var result = await recognizer.RecognizeOnceAsync().ConfigureAwait(false);

                //transcribedMessage = result.Text;
                transcribedMessage = "";
                await ContinuousRecognitionWithFileAsync();

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
                //transcribedMessage = "";
            }
            client.Disconnect(true);
        }
    }

    // Continuous speech recognition.
    public static async Task ContinuousRecognitionWithFileAsync()
    {
        // Transcribing voice message old program
        // Transcribing
        var speechConfig = SpeechConfig.FromSubscription(AzureSubscription, AzureServer);
        speechConfig.SpeechRecognitionLanguage = "sv-SE";
        //await FromFile(speechConfig);

        // Ändra från testPath
        string[] fileArray = Directory.GetFiles(FileFromPath, "*.wav");
        FileFromName = fileArray[0];
        //using var audioConfig = AudioConfig.FromWavFileInput(FileFromName);
        //using var audioConfig = AudioConfig.FromWavFileInput("PathToFile.wav");
        //using var recognizer = new SpeechRecognizer(speechConfig, audioConfig);

        //var result = await recognizer.RecognizeOnceAsync().ConfigureAwait(false);

        //TranscribedMessage = result.Text;

        // <recognitionContinuousWithFile>
        // Creates an instance of a speech config with specified subscription key and service region.
        // Replace with your own subscription key and service region (e.g., "westus").


        var stopRecognition = new TaskCompletionSource<int>();

        // Creates a speech recognizer using file as audio input.
        // Replace with your own audio file name.
        using (var audioConfig = AudioConfig.FromWavFileInput(FileFromName))
        {
            using (var recognizer = new SpeechRecognizer(speechConfig, audioConfig))
            {
                // Subscribes to events.
                recognizer.Recognizing += (s, e) =>
                {
                    Console.WriteLine($"RECOGNIZING: Text={e.Result.Text}");
                    tempMessage = e.Result.Text;
                };
                transcribedMessage += tempMessage;
                recognizer.Recognized += (s, e) =>
                {
                    if (e.Result.Reason == ResultReason.RecognizedSpeech)
                    {
                        Console.WriteLine($"RECOGNIZED: Text={e.Result.Text}");
                        tempMessage = e.Result.Text;
                    }
                    else if (e.Result.Reason == ResultReason.NoMatch)
                    {
                        Console.WriteLine($"NOMATCH: Speech could not be recognized.");
                    }
                    transcribedMessage += tempMessage;
                };

                recognizer.Canceled += (s, e) =>
                {
                    Console.WriteLine($"CANCELED: Reason={e.Reason}");

                    if (e.Reason == CancellationReason.Error)
                    {
                        Console.WriteLine($"CANCELED: ErrorCode={e.ErrorCode}");
                        Console.WriteLine($"CANCELED: ErrorDetails={e.ErrorDetails}");
                        Console.WriteLine($"CANCELED: Did you update the subscription info?");
                    }

                    stopRecognition.TrySetResult(0);
                };

                recognizer.SessionStarted += (s, e) =>
                {
                    Console.WriteLine("\n    Session started event.");
                };

                recognizer.SessionStopped += (s, e) =>
                {
                    Console.WriteLine("\n    Session stopped event.");
                    Console.WriteLine("\nStop recognition.");
                    stopRecognition.TrySetResult(0);
                };

                // Starts continuous recognition. Uses StopContinuousRecognitionAsync() to stop recognition.
                await recognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

                // Waits for completion.
                // Use Task.WaitAny to keep the task rooted.
                Task.WaitAny(new[] { stopRecognition.Task });

                // Stops recognition.
                await recognizer.StopContinuousRecognitionAsync().ConfigureAwait(false);
                Console.WriteLine(transcribedMessage);
            }
        }
    }
    // </recognitionContinuousWithFile>

    static void CheckFolder(string targetPath)
    {
        if (!Directory.Exists(targetPath))
        {
            Directory.CreateDirectory(targetPath);
        }
    }

    static void CheckJson()
    {
        if (!File.Exists(jsonSettings))
        {
            //File.Create(jsonSettings);
            Console.WriteLine("File settings.json don't exist");
            JObject o1 = new JObject(
                new JProperty("FileFromPath", "C:\\AppSpeak\\ljud"),
                new JProperty("moveTo", "Transkriberad"),
                new JProperty("fromName", "Avvikelse"),
                new JProperty("fromEmail", "avvikelse@testboka.net"),
                new JProperty("fromClient", "mail.simply.com"),
                new JProperty("fromPort", 110),
                new JProperty("fromImapPort", 143),
                new JProperty("toClient", "smtp.simply.com"),
                new JProperty("toPort", 587),
                new JProperty("passw", "XXX"),
                new JProperty("toName", "Avvikelse2"),
                new JProperty("toEmail", "avvikelse2@testboka.net"),
                new JProperty("AzureSubscription", "XXX"),
                new JProperty("AzureServer", "westeurope"));
            // write JSON directly to a file
            using (StreamWriter file = File.CreateText(jsonSettings))
            using (JsonTextWriter writer = new JsonTextWriter(file))
            {
                o1.WriteTo(writer);
            }
        }
    }

    static void ReadJson()
    {
        JObject o1 = JObject.Parse(File.ReadAllText(jsonSettings));

        // För test  -----
        //foreach (JProperty property in o1.Properties())
        //{
        //    Console.WriteLine(property.Name + " - " + property.Value);
        //}
        //Console.WriteLine();
        // -------

        FileFromPath = (string)o1["FileFromPath"];
        moveTo = (string)o1["moveTo"];
        fromName = (string)o1["fromName"];
        fromEmail = (string)o1["fromEmail"];
        fromClient = (string)o1["fromClient"];
        fromPort = (int)o1["fromPort"];
        fromImapPort = (int)o1["fromImapPort"];
        toClient = (string)o1["toClient"];
        toPort = (int)o1["toPort"];
        passw = (string)o1["passw"];
        toName = (string)o1["toName"];
        toEmail = (string)o1["toEmail"];
        AzureSubscription = (string)o1["AzureSubscription"];
        AzureServer = (string)o1["AzureServer"];
    }

    public static void MoveToFolderImap()
    {
        using (var client = new ImapClient())
        {
            client.Connect(fromClient, fromImapPort, false);

            client.Authenticate(fromEmail, passw);

            // The Inbox folder is always available on all IMAP servers...
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadWrite);

            // Bara för test -------
            //Console.WriteLine("Total messages: {0}", inbox.Count);
            //Console.WriteLine("Recent messages: {0}", inbox.Recent);
            // -------

            var uids = inbox.Search(SearchQuery.All);

            IMailFolder destination = client.GetFolder(moveTo);
            var uidMap = inbox.MoveTo(uids, destination);
            
            client.Disconnect(true);
        }
    }

    public static void DeleteWav()
    {
        DirectoryInfo di = new DirectoryInfo(FileFromPath);

        foreach (FileInfo file in di.GetFiles())
        {
            file.Delete();
        }
    }

    public static void ErrorLogging(Exception ex)
    {
        string strPath = "C:\\AppSpeak\\Log.txt";
        if (!File.Exists(strPath))
        {
            File.Create(strPath).Dispose();
        }
        using (StreamWriter sw = File.AppendText(strPath))
        {
            sw.WriteLine("=============Error Logging ===========");
            sw.WriteLine("===========Start============= " + DateTime.Now);
            sw.WriteLine("Error Message: " + ex.Message);
            sw.WriteLine("Stack Trace: " + ex.StackTrace);
            sw.WriteLine("===========End============= " + DateTime.Now);
        }
    }

    public static void ReadError()
    {
        string strPath = @"C:\AppSpeak\Log.txt";
        using (StreamReader sr = new StreamReader(strPath))
        {
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
    }

    async static Task Main(string[] args)
    {
        try
        {
            CheckFolder(FileFromPath);
            CheckJson();
            ReadJson();
            // Transcribing
            //var speechConfig = SpeechConfig.FromSubscription(AzureSubscription, AzureServer);
            //speechConfig.SpeechRecognitionLanguage = "sv-SE";
            //await FromFile(speechConfig);
            await FromFile();

            MoveToFolderImap();

            // Deleting all files in folder ljud
            DeleteWav();
        }
        catch (Exception ex)
        {
            ErrorLogging(ex);
            ReadError();
        }   
    }
 
}



