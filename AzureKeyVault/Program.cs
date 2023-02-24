// See https://aka.ms/new-console-template for more information

using System.Text;
using AzureKeyVault.Utilities;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using File = System.IO.File;


namespace AzureKeyVault;

internal class Program
{
    public static async Task Main(string[] args)
    {
        Console.WriteLine("Azure Test");

        //  These values come from your Azure account
        string tenantId = "";
        string clientId ="";
        string clientSecret = "";
        
        const string email = "";
        const string destFoldername = "";
        
        // Get Token From Azure
        Task<AuthenticationResult> task = tenantId.GetToken(clientId, clientSecret,"v2.0");
        task.Wait();
        if (task.IsFaulted)
        {
            throw new Exception("Access token not found.");
        }
        AuthenticationResult token = task.Result;
        Console.WriteLine($"Token: {token.AccessToken}\tAccount: {token.Account}\tToken ID: {token.IdToken}");

        // Read Email's
        try
        {
            GraphServiceClient graphClientCs = tenantId.ClientSecret(clientId, clientSecret);
            
            // Get Destination Folder
            IUserMailFoldersCollectionPage? folders = await graphClientCs.Users[email]
                .MailFolders
                .Request()
                .Filter($"displayName eq '{destFoldername}'")
                .GetAsync();
            MailFolder? destFolder = folders[0];
            
            IUserMailFoldersCollectionPage? test = await graphClientCs.Users[email]
                .MailFolders
                .Request()
                .Filter($"displayName eq 'FcaFasciaTest'")
                .GetAsync();
            MailFolder? testFolder = test[0];
            
            var test1 = await graphClientCs.Users[email]
                .MailFolders[testFolder.Id]
                .Messages
                .Request()
                .Filter("IsRead eq false and HasAttachments eq true")
                .Select("sender,subject,body,hasattachments,isread")
                .GetAsync();
            
            // Get Mail Messages
            IMailFolderMessagesCollectionPage? mailmsg = await graphClientCs
                .Users[email]
                .MailFolders[testFolder.Id]
                .Messages
                .Request()
                .Filter("IsRead eq false and HasAttachments eq true")
                .Select("sender,subject,body,hasattachments,isread")
                .GetAsync();
            
            // Process the attachments
            foreach (Message? msg in mailmsg)
            {
                Console.WriteLine($"{msg.Sender.EmailAddress.Address} - {msg.Subject} - {msg.HasAttachments}");
                if (msg.HasAttachments is null || !msg.HasAttachments.Value) continue;
        
                IMessageAttachmentsCollectionPage? attachments = await graphClientCs.Users[email].Messages[msg.Id].Attachments
                    .Request()
                    .GetAsync();
        
                foreach (Attachment? attach in attachments)
                {
                    if (attach is not FileAttachment fa) continue;
                    string extension = Path.GetExtension(attach.Name);
                    //if (!extension.Equals(".csv.")) continue;
            
                    string encodedString = Convert.ToBase64String(fa.ContentBytes);
                    byte[] data = Convert.FromBase64String(encodedString);

                    if (!extension.Equals(".csv."))
                    {
                        await File.WriteAllBytesAsync($"C:\\Temp\\{attach.Name}", data);
                    }
                    else
                    {
                        string decodedString = Encoding.UTF8.GetString(data);
                        Console.WriteLine($"FileName: {attach.Name}");
                        await File.WriteAllBytesAsync($"C:\\Temp\\{attach.Name}", data);
                        
                        // Write Contents to a file or should return <fileName, byte[]>
                        //await System.IO.File.WriteAllTextAsync($@"C:\Temp\{attach.Name}", decodedString);
                    }
                }
                
                Message? t = await graphClientCs.Users[email].Messages[msg.Id]
                    .Request()
                    .Select("IsRead")
                    .UpdateAsync(new Message {IsRead = true});
                    
                // Move the Mail Message to the Process Folder
                await graphClientCs.Users[email].Messages[msg.Id]
                    .Move(destFolder.Id)
                    .Request()
                    .PostAsync();
               
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }

        Console.WriteLine("Done!");
    }
}