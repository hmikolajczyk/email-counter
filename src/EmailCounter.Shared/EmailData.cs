namespace EmailCounter.Shared.Models
{
    public class EmailData
    {
        public string? Subject { get; set; }
        public DateTime ReceivedTime { get; set; }
        public required string Sender { get; set; }
        public required string ConversationID { get; set; }
        public required string ConversationTopic { get; set; }
    }
}