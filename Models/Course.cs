namespace LearnX.Models
{
    public class Course
    {
        public int CourseId { get; set; }
        public string CourseName { get; set; }
        public string CourseDescription { get; set; }
        public string CourseImageURL { get; set; }
        public string CourseRating { get; set; }

        public int CourseDuration { get; set; }
        public string CourseDifficulty { get; set; }
        public int EnrolledStudentsCount { get; set; }
        public string CourseOverview { get; set; }
        public string YoutubeLink { get; set; }
        public string NotesLink { get; set; }

    }
}
