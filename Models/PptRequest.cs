namespace ppt_mapper.Models
{
    public class PptRequest
    {
        public string? Title { get; set; }
        public bool TitleIsBold { get; set; }

        public string? Date { get; set; }
        public bool DateIsBold { get; set; }

        public string? SprintNumber { get; set; }
        public bool SprintNumberIsBold { get; set; }

        public List<string>? Points { get; set; }

        public List<string>? Page1Points { get; set; }
        public List<bool>? Page1PointIsBulletPoint { get; set; } = CreateFlags(11, true);
        public List<bool>? Page1PointIsBold { get; set; } = CreateFlags(11, false);

        public List<string>? Page2Points { get; set; }
        public List<bool>? Page2PointIsBulletPoint { get; set; } = CreateFlags(11, true);
        public List<bool>? Page2PointIsBold { get; set; } = CreateFlags(11, false);

        public List<string>? Page3Points { get; set; }
        public List<bool>? Page3PointIsBulletPoint { get; set; } = CreateFlags(11, true);
        public List<bool>? Page3PointIsBold { get; set; } = CreateFlags(11, false);

        public List<string>? Page4Points { get; set; }
        public List<bool>? Page4PointIsBold { get; set; } = CreateFlags(45, false);

        public List<string>? Page5Points { get; set; }
        public List<bool>? Page5PointIsBold { get; set; } = CreateFlags(35, false);

        public List<string>? Page6Points { get; set; }
        public List<bool>? Page6PointIsBulletPoint { get; set; } = CreateFlags(11, true);
        public List<bool>? Page6PointIsBold { get; set; } = CreateFlags(11, false);

        public List<string>? Page7Points { get; set; }
        public List<bool>? Page7PointIsBulletPoint { get; set; } = CreateFlags(11, true);
        public List<bool>? Page7PointIsBold { get; set; } = CreateFlags(11, false);

        public List<string>? NumericValues { get; set; }
        public List<bool>? NumericValueIsBold { get; set; } = CreateFlags(20, false);

        public List<string>? PercentageValues { get; set; }
        public List<bool>? PercentageValueIsBold { get; set; } = CreateFlags(8, false);

        public string? TotalStoriesCommitted { get; set; }
        public bool TotalStoriesCommittedIsBulletPoint { get; set; } = true;
        public bool TotalStoriesCommittedIsBold { get; set; }

        public string? TotalStoryPoints { get; set; }
        public bool TotalStoryPointsIsBulletPoint { get; set; } = true;
        public bool TotalStoryPointsIsBold { get; set; }

        public string? NewUserStories { get; set; }
        public bool NewUserStoriesIsBulletPoint { get; set; } = true;
        public bool NewUserStoriesIsBold { get; set; }

        public string? SpilloverStories { get; set; }
        public bool SpilloverStoriesIsBulletPoint { get; set; } = true;
        public bool SpilloverStoriesIsBold { get; set; }

        public string? AgileCeremonyProcessItem { get; set; }
        public bool AgileCeremonyProcessItemIsBulletPoint { get; set; } = true;
        public bool AgileCeremonyProcessItemIsBold { get; set; }

        public List<IFormFile>? Images { get; set; }

        private static List<bool> CreateFlags(int count, bool defaultValue)
        {
            return Enumerable.Repeat(defaultValue, count).ToList();
        }
    }
}
