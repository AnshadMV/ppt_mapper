namespace ppt_mapper.Models
{
    public class PptRequest
    {
        public string Title { get; set; }
        public string Date { get; set; }
        public string SprintNumber { get; set; }
        public List<string> Points { get; set; }
        public List<IFormFile> Images { get; set; }
    }
}
