using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;

namespace XYZCorp.GraphService
{
    public class EmailService
    {
        public static string DownloadFolder = ConfigurationManager.AppSettings["DownloadFolder"];
        private GraphServiceClient _graphClient;

        public EmailService()
        {
            _graphClient = CredentialManager.GetGraphClientWithClientSecret();

            //_graphClient = CredentialManager.GetGraphClientWithInteractive();

            // _graphClient = CredentialManager.GetGraphClientWithUserIDPwd();
        }

        /// <summary>
        /// Read the emails for the input emailId
        /// </summary>
        /// <param name="emailId"></param>
        /// <param name="fromDate"></param>
        /// <returns></returns>
        public async Task<List<EmailMessage>> GetEmails(string emailId, DateTime fromDate)
        {
            var user = await _graphClient.Users.Request().Filter($"mail eq '{emailId}'").GetAsync();
            if (user.Count == 0) throw new Exception($"Invalid Email ID: {emailId}");

            var currentUser = await _graphClient.Users[emailId].Request().GetAsync();

            return await GetUserEmails(currentUser.Id, fromDate);
        }

        public async Task GetAttachment(AttachmentDto attachment)
        {
            var file = await _graphClient
                                    .Users[attachment.UserId]
                                    .Messages[attachment.MessageId]
                                    .Attachments[attachment.AttachmentId]
                                    .Request().GetAsync();

            if (file.ODataType.Equals("#microsoft.graph.fileAttachment", StringComparison.OrdinalIgnoreCase))
            {
                FileAttachment att = (FileAttachment)file;
                string path = String.Concat(DownloadFolder, file.Name);
                System.IO.File.WriteAllBytes(path, att.ContentBytes);
            }
        }

        /// <summary>
        /// Use this method to get the emails for the logged in user's account
        /// </summary>
        /// <param name="fromDate"></param>
        /// <returns></returns>
        public async Task<List<EmailMessage>> GetMyEmails(DateTime fromDate)
        {
            var currentUser = await _graphClient.Me.Request().GetAsync();

            return await GetUserEmails(currentUser.Id, fromDate);
        }

        private async Task<List<EmailMessage>> GetUserEmails(string userId, DateTime fromDate)
        {
            List<EmailMessage> emails = new List<EmailMessage>();

            string isoFormat = String.Concat(fromDate.ToString("yyyy-MM-dd"), "T00:00:00Z");
            string dateFilter = $"ReceivedDateTime ge {isoFormat}";
            
            var messages = await _graphClient.Users[userId].Messages.Request().Filter(dateFilter).GetAsync();

            foreach (var message in messages)
            {
                List<AttachmentDto> attachments = new List<AttachmentDto>();

                if ((bool)message.HasAttachments)
                {
                    var msgAttachments = await _graphClient.Users[userId].Messages[message.Id].Attachments.Request().GetAsync();

                    foreach (var attachment in msgAttachments)
                    {
                        attachments.Add(new AttachmentDto()
                        {
                            UserId = userId,
                            MessageId = message.Id,
                            AttachmentId = attachment.Id,
                            Name = attachment.Name,
                            ContentType = attachment.ContentType
                        });
                    }
                }

                emails.Add(new EmailMessage()
                {
                    MessageId = message.Id,
                    SenderEmail = message.From.EmailAddress.Address,
                    Subject = message.Subject,
                    ReceivedDate = message.ReceivedDateTime.Value.Date,
                    Attachments = attachments
                });
            }

            return emails;
        }        
    }
}