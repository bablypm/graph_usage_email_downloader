using System;
using System.Collections.Generic;

namespace XYZCorp.GraphService
{
    public class EmailMessage
    {
        public string MessageId { get; set; }
        public string SenderEmail { get; set; }
        public string Subject { get; set; }
        public DateTime ReceivedDate { get; set; }
        public List<AttachmentDto> Attachments { get; set; }
    }
}