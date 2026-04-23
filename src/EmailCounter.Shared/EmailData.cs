namespace EmailCounter.Shared
{
    public class EmailData
    {
        public required string Subject { get; set; }
        public DateTime ReceivedTime { get; set; }
        public required string Sender { get; set; }
        public required string ConversationID { get; set; }
        public required string ConversationTopic { get; set; }
    }
}