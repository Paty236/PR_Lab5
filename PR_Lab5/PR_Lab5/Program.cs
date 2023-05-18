using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Net.Mail;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using MimeKit;
using MailKit.Net.Pop3;
using System;
using MimeKit.Text;

string popServer = "pop.gmail.com";
int popPort = 995;
string imapServer = "imap.gmail.com";
int imapPort = 993;
string username = "reitmanpatricia6@gmail.com";
string password = "temcdpcizmargrox";

bool exit = false;
while (!exit)
{
    Console.WriteLine("\nMenu:");
    Console.WriteLine("1. List emails in mailbox using POP3 protocol");
    Console.WriteLine("2. List emails in mailbox using IMAP protocol");
    Console.WriteLine("3. Download an email with attachments");
    Console.WriteLine("4. Send an email with text only");
    Console.WriteLine("5. Send an email with attachments");
    Console.WriteLine("6. Include subject and reply-to details when sending email");
    Console.WriteLine("0. Exit");
    Console.Write("\nEnter desired option: ");
    var option = Console.ReadLine();

    switch (option)
    {
        case "1":
            ListEmailsPop3();
            break;

        case "2":
            ListEmailsImap();
            break;

        case "3":
            DownloadEmail();
            break;

        case "4":
            SendTextEmail();
            break;

        case "5":
            SendEmailWithAttachments();
            break;

        case "6":
            SendEmailWithDetails();
            break;

        case "0":
            exit = true;
            break;

        default:
            Console.WriteLine("Invalid option. Please try again.");
            break;
    }
}

void ListEmailsPop3()
{
    // Connect to POP3 server
    using var tcpClient = new TcpClient(popServer, popPort);
    using var sslStream = new SslStream(tcpClient.GetStream());
    sslStream.AuthenticateAsClient(popServer);

    // Set up reader and writer for communication with server
    using var reader = new StreamReader(sslStream, Encoding.ASCII);
    using var writer = new StreamWriter(sslStream, Encoding.ASCII);
    string response = reader.ReadLine() ?? "";

    // Login to email account
    writer.WriteLine("USER " + username);
    writer.Flush();
    response = reader.ReadLine() ?? "";

    writer.WriteLine("PASS " + password);
    writer.Flush();
    response = reader.ReadLine() ?? "";

    // Retrieve list of messages in mailbox
    writer.WriteLine("LIST");
    writer.Flush();
    response = reader.ReadLine() ?? "";

    // Print list of messages to console
    while ((response = reader.ReadLine() ?? "") != "." && response != null)
    {
        Console.WriteLine(response);
    }

    // Logout of email account
    writer.WriteLine("QUIT");
    writer.Flush();
    response = reader.ReadLine() ?? "";

    Console.WriteLine("POP3 email list retrieved successfully.");
}

void ListEmailsImap()
{
    // Connect to IMAP server
    using var client = new ImapClient();
    client.Connect(imapServer, imapPort, SecureSocketOptions.SslOnConnect);
    client.Authenticate(username, password);

    // Open inbox and print message count and subject lines to console
    var inbox = client.Inbox;
    inbox.Open(FolderAccess.ReadOnly);

    Console.WriteLine("Total Messages: {0}", inbox.Count);

    for (int i = 0; i < inbox.Count; i++)
    {
        var message = inbox.GetMessage(i);
        Console.WriteLine("Subject: {0}", message.Subject);
        Console.WriteLine("From: {0}", message.From);
        Console.WriteLine();
    }

    Console.WriteLine("IMAP email list retrieved successfully.");
}

void DownloadEmail()
{
    // Define download path for email attachments
    string downloadPath = "\\Users\\patyreitman\\Desktop\\";

    // Connect to POP3 server and authenticate
    using (var pop3Client = new Pop3Client())
    {
        pop3Client.Connect(popServer, popPort, true);
        pop3Client.Authenticate(username, password);

        // Loop through all emails in mailbox
        for (int i = 0; i < pop3Client.Count; i++)
        {
            var message = pop3Client.GetMessage(i);

            // Loop through all attachments in email
            foreach (var attachment in message.Attachments)
            {
                // Get file name from attachment
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                // Create file and write attachment content to it
                using (var stream = File.Create(downloadPath + fileName))
                {
                    if (attachment is MessagePart)
                    {
                        var part = (MessagePart)attachment;
                        part.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)attachment;
                        part.Content.DecodeTo(stream);
                    }
                }

                Console.WriteLine("Downloaded file: " + fileName);
            }
        }
    }
}

void SendTextEmail()
{
    // Get email recipient address from user input
    string toEmail;
    do
    {
        Console.Write("Enter email recipient address: ");
        toEmail = Console.ReadLine() ?? "";

        if (!Regex.IsMatch(toEmail, @"^[^@\s]+@([a-zA-Z0-9]+\.)+[a-zA-Z]{2,}$"))
        {
            Console.WriteLine("Invalid email address. Please enter a valid email address.");
            toEmail = "";
        }
    } while (toEmail == null);

    // Get email subject and body from user input
    Console.Write("Enter email subject: ");
    string subject = Console.ReadLine() ?? "";

    Console.Write("Enter email body: ");
    string body = Console.ReadLine() ?? "";

    // Connect to SMTP server and authenticate
    var smtpClient = new SmtpClient("smtp.gmail.com", 587);
    smtpClient.EnableSsl = true;
    smtpClient.UseDefaultCredentials = false;
    smtpClient.Credentials = new NetworkCredential(username, password);

    // Create email message
    var message = new MailMessage(username, toEmail);
    message.Subject = subject;
    message.Body = body;

    // Send email message
    try
    {
        smtpClient.Send(message);
        Console.WriteLine("Email sent successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error sending email: " + ex.Message);
    }
}

void SendEmailWithAttachments()
{
    string toEmail;
    do
    {
        Console.Write("Enter recipient email address: ");
        toEmail = Console.ReadLine() ?? "";

        if (!Regex.IsMatch(toEmail, @"^[^@\s]+@([a-zA-Z0-9]+\.)+[a-zA-Z]{2,}$"))
        {
            Console.WriteLine("Invalid email address. Please enter a valid email address.");
            toEmail = "";
        }
    } while (toEmail == null);

    Console.Write("Enter email subject: ");
    string subject = Console.ReadLine() ?? "";

    Console.Write("Enter email message: ");
    string body = Console.ReadLine() ?? "";

    Console.Write("Enter path to attachment file: ");
    string attachmentPath = Console.ReadLine() ?? "";

    Attachment attachment = new Attachment(attachmentPath);

    MailMessage mailMessage = new MailMessage();
    mailMessage.From = new MailAddress(username);
    mailMessage.To.Add(toEmail);
    mailMessage.Subject = subject;
    mailMessage.Body = body;
    mailMessage.Attachments.Add(attachment);

    var smtpClient = new SmtpClient("smtp.gmail.com", 587);
    smtpClient.EnableSsl = true;
    smtpClient.UseDefaultCredentials = false;
    smtpClient.Credentials = new NetworkCredential(username, password);

    try
    {
        smtpClient.Send(mailMessage);
        Console.WriteLine("Email sent successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error sending email: " + ex.Message);
    }
}

void SendEmailWithDetails()
{
    string toEmail;
    do
    {
        Console.Write("Enter recipient email address: ");
        toEmail = Console.ReadLine() ?? "";

        if (!Regex.IsMatch(toEmail, @"^[^@\s]+@([a-zA-Z0-9]+\.)+[a-zA-Z]{2,}$"))
        {
            Console.WriteLine("Invalid email address. Please enter a valid email address.");
            toEmail = "";
        }
    } while (toEmail == null);

    string replyTo;
    do
    {
        Console.WriteLine("Enter email address for reply:");
        replyTo = Console.ReadLine() ?? "";

        if (!Regex.IsMatch(replyTo, @"^[^@\s]+@([a-zA-Z0-9]+\.)+[a-zA-Z]{2,}$"))
        {
            Console.WriteLine("Invalid email address. Please enter a valid email address.");
            replyTo = "";
        }
    } while (replyTo == null);

    Console.Write("Enter email subject: ");
    string subject = Console.ReadLine() ?? "";

    Console.Write("Enter email message: ");
    string body = Console.ReadLine() ?? "";

    MailMessage message = new MailMessage();
    message.To.Add(toEmail);
    message.Subject = subject;
    message.Body = body;
    message.ReplyToList.Add(replyTo);

    var smtpClient = new SmtpClient("smtp.gmail.com", 587);
    smtpClient.EnableSsl = true;
    smtpClient.UseDefaultCredentials = false;
    smtpClient.Credentials = new NetworkCredential(username, password);

    try
    {
        smtpClient.Send(message);
        Console.WriteLine("Email sent successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error sending email: " + ex.Message);
    }
}
