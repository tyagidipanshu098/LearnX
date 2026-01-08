using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using LearnX.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.BlazorIdentity.Pages.Manage;
using System.Diagnostics;
using System.Text.Json;

namespace LearnX.Controllers
{
    public class HomeController : Controller
    {
  
        public IActionResult Index()
        {
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                var courses = GetCourses();
                return View(courses);
            }
            return RedirectToAction("Login");
        }


        private List<Course> GetCourses()
        {
            return new List<Course>(){
                new Course() {
                    CourseId = 1,
                    CourseName = ".NET Core MVC",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/NET.png",
                    CourseRating = "4.6",
                    CourseDuration = 10,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 55,
                    CourseOverview = "Master the architecture of modern web applications using the Model-View-Controller (MVC) pattern. This course covers the latest features of .NET 8/9+, teaching you how to build scalable, high-performance web systems. You will learn about dependency injection, middleware configuration, secure authentication, and the Razor View engine to create dynamic, data-driven user interfaces that follow industry best practices for clean code and separation of concerns.",
                    YoutubeLink = "https://youtube.com/playlist?list=PLp_RsiLZjwQQ7CxVhnM4G8i5veEClNPfX&si=j0sEJ2Co2Ryv_sRC",
                    NotesLink = "",
                },
                new Course() {
                    CourseId = 2,
                    CourseName = "C#",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/CSharp.png",
                    CourseRating = "4.5",
                    CourseDuration = 11,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 40,
                    CourseOverview = "Dive into the world’s most versatile programming language. This course takes you from fundamental syntax to advanced object-oriented programming (OOP) concepts. You will explore modern C# features including LINQ for data manipulation, asynchronous programming with Async/Await for responsive apps, and memory management. Whether you're aiming for web, desktop, or game development, this course provides the bedrock logic required for professional software engineering.",

                },
                new Course() {
                    CourseId = 3,
                    CourseName = "Debugging",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/Debugging.png",
                    CourseRating = "4.8",
                    CourseDuration = 14,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 59,
                    CourseOverview = "The difference between a junior and a senior developer is the ability to solve problems quickly. This course focuses on the \"Art of the Fix.\" You will learn to use the Visual Studio debugger like a pro, mastering breakpoints, watch windows, and call stacks. Beyond the tools, you’ll develop the mindset needed to isolate bugs in complex systems, handle exceptions gracefully, and use logging frameworks to diagnose issues in production environments.",

                },
                new Course() {
                    CourseId = 4,
                    CourseName = "SQL",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/SQL1.png",
                    CourseRating = "4.9",
                    CourseDuration = 13,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 35,
                    CourseOverview = "Data is the heart of every application. In this course, you will learn how to communicate with relational databases using SQL. We cover everything from basic SELECT statements and complex JOIN operations to writing efficient stored procedures and triggers. You will learn how to structure queries for maximum performance and how to ensure data integrity, preparing you to handle the back-end logic of any modern enterprise application.",

                },
                new Course() {
                    CourseId = 5,
                    CourseName = "JQuery",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/JQuery.png",
                    CourseRating = "4.2",
                    CourseDuration = 17,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 74,
                    CourseOverview = "Learn how to simplify client-side scripting and DOM manipulation. Despite the rise of modern frameworks, jQuery remains a vital tool for maintaining legacy systems and rapidly prototyping interactive elements. This course teaches you how to handle events, create smooth animations, and perform AJAX requests to update page content without a refresh. You will learn to write \"less code\" while doing \"more\" in the browser.",

                },
                new Course() {
                    CourseId = 6,
                    CourseName = "Database",
                    CourseDescription = "Beginner course",
                    CourseImageURL = "/Images/Database.png",
                    CourseRating = "4.5",
                    CourseDuration = 15,
                    CourseDifficulty = "Beginner",
                    EnrolledStudentsCount = 50,
                    CourseOverview = "Building a database is easy; designing a good one is a science. This course covers the lifecycle of data management, from Entity-Relationship Diagrams (ERDs) and Normalization to indexing strategies and database security. You will learn the differences between Relational (SQL) and Non-Relational (NoSQL) systems, ensuring you can choose and build the right storage architecture for any project’s specific needs.",

                },
            };
        }


        public IActionResult CourseView(int id)
        { 
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                var selectedCourse = GetCourses().FirstOrDefault(x => x.CourseId == id);
                if (selectedCourse == null) return NotFound();
                var viewModel = new CourseFeedbackViewModel
                {
                    CourseData = selectedCourse,
                };
                return View(viewModel);
            }
            return RedirectToAction("Login");
            
        }


        //--------Excel submission

        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
        }


       

        //-------Login


        public IActionResult Login()
        {
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                return RedirectToAction("Index");
            }
            return View();
        }

        [HttpPost]
        public IActionResult Login(Student student)
        {
            string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");
            bool isMatch = false;

            if (System.IO.File.Exists(filePath))
            {
                // Use a FileStream with FileShare.Read to prevent "File in use" errors 
                // if the Excel file is open while someone tries to login
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(fs))
                {
                    var worksheet = workbook.Worksheet(1);
                    var usedRange = worksheet.RangeUsed();

                    if (usedRange != null)
                    {
                        var rows = usedRange.RowsUsed().Skip(1);
                        foreach (var row in rows)
                        {
                            // Use ?.Trim() to handle accidental whitespace in Excel cells
                            string nameFromExcel = row.Cell(1).GetValue<string>()?.Trim();
                            string passwordFromExcel = row.Cell(4).GetValue<string>()?.Trim();

                            if (nameFromExcel == student.StudentName && passwordFromExcel == student.StudentPassword)
                            {
                                isMatch = true;
                                break;
                            }
                        }
                    }
                }
            }

            if (isMatch)
            {
                HttpContext.Session.SetString("UserSession", student.StudentName);
                HttpContext.Session.SetString("UserName", student.StudentName);
                TempData["LoggedIn"] = "Logged In successfully!";
                return RedirectToAction("Index");
            }

            ViewBag.Message = "Invalid name or password.";
            return View();
        }


        public IActionResult Logout()
        {
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                HttpContext.Session.Remove("UserSession");
                TempData["LoggedOut"] = "Logged out successfully!";
                return RedirectToAction("Login");
            }
            return View();
        }

        public IActionResult Register()
        {

            return View();
        }

        [HttpPost]
        public IActionResult Register(Student student)
        { 

            if (!ModelState.IsValid)
                return View("Index");

            if (student.StudentPassword != student.StudentConfirmPassword)
            {
                ViewBag.PasswordNotMatched = "Password not matched.";
                return RedirectToAction("Register");
            }
            string folderPath = Path.Combine(_env.ContentRootPath, "Excel");
            Directory.CreateDirectory(folderPath);

            string filePath = Path.Combine(folderPath, "RegisteredStudents.xlsx");

            using var workbook = System.IO.File.Exists(filePath)
                ? new XLWorkbook(filePath)
                : new XLWorkbook();

            var worksheet = workbook.Worksheets.FirstOrDefault()
                            ?? workbook.AddWorksheet("Submissions");

            // Add header row only once
            if (worksheet.LastRowUsed() == null)
            {
                worksheet.Cell(1, 1).Value = "Student Name";
                worksheet.Cell(1, 2).Value = "Student Email";
                worksheet.Cell(1, 3).Value = "Student Phone";
                worksheet.Cell(1, 4).Value = "Student Password";
                worksheet.Cell(1, 5).Value = "Registration date";
            }

            int nextRow = (worksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

            worksheet.Cell(nextRow, 1).Value = student.StudentName;
            worksheet.Cell(nextRow, 2).Value = student.StudentEmail;
            worksheet.Cell(nextRow, 3).Value = student.StudentPhone;
            worksheet.Cell(nextRow, 4).Value = student.StudentPassword;
            worksheet.Cell(nextRow, 5).Value = DateTime.Now;
      

            workbook.SaveAs(filePath);

            TempData["Registered"] = "Registered successfully!";
            HttpContext.Session.SetString("UserEmail", student.StudentEmail);
            return RedirectToAction("Login");
        }


        public IActionResult Forget()
        {
            return View(new Student());
        }


        [HttpPost]
        public IActionResult Forget(Student studentInput)
        {
            // Re-initialize found to false for each search attempt
            ViewBag.found = false;

            if (string.IsNullOrEmpty(studentInput.StudentName))
            {
                TempData["Error"] = "Please provide a name.";
                return View(studentInput); // Pass model back
            }

            string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");

            if (System.IO.File.Exists(filePath))
            {
                // Use FileShare.Read to allow multiple reads without locking
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(fs))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString().Equals(studentInput.StudentName, StringComparison.OrdinalIgnoreCase))
                        {
                            ViewBag.found = true;
                            break;
                        }
                    }
                }
            }
            // Return the model so @Model.StudentName is still available in the hidden field
            return View(studentInput);
        }


        [HttpPost]
        public IActionResult ChangePassword(string NewPassword, string StudentName)
        {
                string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(fs))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        if (row.Cell(1).GetString() == StudentName)
                        {
                            row.Cell(4).Value = NewPassword;
                            workbook.Save();

                        TempData["PasswordUpdate"] = "Password changed successfully!";
                            break;
                        }
                    }
                }
                return RedirectToAction("Login");
        }


        [HttpPost]
        public IActionResult Feedback(CourseFeedbackViewModel cf)
        {
            string studentName = HttpContext.Session.GetString("UserName") ?? "Anonymous";

            try
            {
                string folderPath = Path.Combine(_env.ContentRootPath, "Excel");
                Directory.CreateDirectory(folderPath);
                string filePath = Path.Combine(folderPath, "Feedback.xlsx");

                // Use a single FileStream for both reading and writing
                using (var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (var workbook = fs.Length > 0 ? new XLWorkbook(fs) : new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.FirstOrDefault() ?? workbook.AddWorksheet("Submissions");

                        if (worksheet.LastRowUsed() == null)
                        {
                            worksheet.Cell(1, 1).Value = "Student Name";
                            worksheet.Cell(1, 2).Value = "Course Id";
                            worksheet.Cell(1, 3).Value = "Course Name";
                            worksheet.Cell(1, 4).Value = "Feedback date";
                            worksheet.Cell(1, 5).Value = "Feedback";
                        }

                        int nextRow = (worksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;
                        worksheet.Cell(nextRow, 1).Value = studentName;
                        worksheet.Cell(nextRow, 2).Value = cf.CourseData.CourseId;
                        worksheet.Cell(nextRow, 3).Value = cf.CourseData.CourseName;
                        worksheet.Cell(nextRow, 4).Value = DateTime.Now;
                        worksheet.Cell(nextRow, 5).Value = cf.FeedbackData.UserFeedback;

                        // CRITICAL: Save to the stream, not the file path, while fs is open
                        workbook.SaveAs(fs);
                    }
                }
                TempData["FeedbackSubmitted"] = "Feedback submitted successfully!";
            }
            catch (IOException)
            {
                // This catch now only triggers if another app (like Excel) has a hard lock on the file
                TempData["Error"] = "The feedback file is currently locked by another program. Please close it and try again.";
            }

            return RedirectToAction("CourseView", new { id = cf.CourseData.CourseId });
        }


        public IActionResult Cart()
        {
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                var cartJson = HttpContext.Session.GetString("Cart");
                var cartItems = string.IsNullOrEmpty(cartJson)
                                ? new List<CartItem>()
                                : JsonSerializer.Deserialize<List<CartItem>>(cartJson);

                return View(cartItems);
            }
            return RedirectToAction("Login");
            
        }


        [HttpPost]
        public IActionResult AddToCart(int id)
        {
            // 1. Fetch the course data
            var course = GetCourses().FirstOrDefault(x => x.CourseId == id);
            if (course == null) return NotFound();

            // 2. Retrieve existing cart string from Session
            string cartJson = HttpContext.Session.GetString("Cart");

            List<CartItem> cart;
            if (string.IsNullOrEmpty(cartJson))
            {
                // If it's the first item, initialize a new list
                cart = new List<CartItem>();
            }
            else
            {
                // Otherwise, turn the JSON string back into a C# List
                cart = JsonSerializer.Deserialize<List<CartItem>>(cartJson);
            }

            // 3. Add the item (prevent duplicates)
            if (!cart.Any(x => x.Course.CourseId == id))
            {
                cart.Add(new CartItem { Course = course, Quantity = 1 });

                // 4. SET THE SESSION HERE
                // Convert the updated list back to a string and save it
                string updatedCartJson = JsonSerializer.Serialize(cart);
                HttpContext.Session.SetString("Cart", updatedCartJson);

                TempData["CartMessage"] = $"{course.CourseName} is added to cart!";
            }

            return RedirectToAction("Index");
        }

        public IActionResult RemoveFromCart(int id)
        {
            var cartJson = HttpContext.Session.GetString("Cart");
            if (!string.IsNullOrEmpty(cartJson))
            {
                var cart = JsonSerializer.Deserialize<List<CartItem>>(cartJson);
                var itemToRemove = cart.FirstOrDefault(x => x.Course.CourseId == id);

                if (itemToRemove != null)
                {
                    cart.Remove(itemToRemove);
                    if (cart.Count > 0)
                        HttpContext.Session.SetString("Cart", JsonSerializer.Serialize(cart));
                    else
                        HttpContext.Session.Remove("Cart");

                    TempData["CartMessage"] = "Item removed from cart.";
                }
            }

            // SMART REDIRECT: Redirects to the page the user was just on
            string referer = Request.Headers["Referer"].ToString();
            if (!string.IsNullOrEmpty(referer))
            {
                return Redirect(referer);
            }

            // Fallback if referer is missing
            return RedirectToAction("Index");
        }


        public IActionResult Profile()
        {
            if (HttpContext.Session.GetString("UserSession") != null)
            {
                string userName = HttpContext.Session.GetString("UserName");
                if (string.IsNullOrEmpty(userName)) return RedirectToAction("Login");

                string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");

                // Initialize with a new object to prevent "Model is null" errors in the View
                Student studentProfile = new Student();

                if (System.IO.File.Exists(filePath))
                {
                    using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var workbook = new XLWorkbook(fs))
                    {
                        var row = workbook.Worksheet(1).RowsUsed().Skip(1)
                                    .FirstOrDefault(r => r.Cell(1).GetValue<string>() == userName);

                        if (row != null)
                        {
                            studentProfile.StudentName = row.Cell(1).GetValue<string>();
                            studentProfile.StudentEmail = row.Cell(2).GetValue<string>();
                            studentProfile.StudentPhone = row.Cell(3).GetValue<string>();
                        }
                    }
                }

                return View(studentProfile);
            }
            return RedirectToAction("Login");
             // studentProfile is now guaranteed to be an instance
        }


        [HttpPost]
        public IActionResult UpdateProfile(Student model)
        {
            // 1. Get the current user session
            string userName = HttpContext.Session.GetString("UserName");
            if (string.IsNullOrEmpty(userName)) return RedirectToAction("Login");

            string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");

            if (System.IO.File.Exists(filePath))
            {
                // Use FileShare.ReadWrite to prevent locking issues in 2026 environments
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(fs))
                {
                    var worksheet = workbook.Worksheet(1);
                    // Find the specific row for this user
                    var row = worksheet.RowsUsed().Skip(1)
                                .FirstOrDefault(r => r.Cell(1).GetValue<string>() == userName);

                    if (row != null)
                    {
                        // Update cells (Column 2 = Email, Column 3 = Phone, etc.)
                        row.Cell(2).Value = model.StudentEmail;
                        row.Cell(3).Value = model.StudentPhone;

                        workbook.Save();
                        TempData["ProfileUpdate"] = "Profile updated successfully!";
                    }
                }
            }

            return RedirectToAction("Profile");
        }




        [HttpPost]
        public IActionResult UpdatePassword(string currentPassword, string newPassword, string confirmPassword, string studentName)
        {
            var filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");

            // 1. Validation: Passwords mismatch
            if (newPassword != confirmPassword)
            {
                TempData["PasswordUnmatched"] = "Confirm password do not match.";
                return RedirectToAction("Profile");
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed().Skip(1);
                bool found = false;

                foreach (var row in rows)
                {
                    // Check Username AND Current Password
                    if (row.Cell(1).GetString() == studentName && row.Cell(4).GetString() == currentPassword)
                    {
                        row.Cell(4).Value = newPassword;
                        found = true;
                        break;
                    }
                }

                if (found)
                {
                    workbook.Save();
                    TempData["PasswordUpdate"] = "Password changed successfully!";
                    return RedirectToAction("Profile");
                }
                else
                {
                    TempData["PasswordUnmatched"] = "Incorrect current password.";
                    return RedirectToAction("Profile");
                }
            }
        }

        // Helper method to keep code clean and prevent null models
        private Student GetStudentFromExcel(string userName)
        {
            string filePath = Path.Combine(_env.ContentRootPath, "Excel", "RegisteredStudents.xlsx");
            Student student = new Student();

            if (System.IO.File.Exists(filePath))
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(fs))
                {
                    var row = workbook.Worksheet(1).RowsUsed().Skip(1)
                                .FirstOrDefault(r => r.Cell(1).GetValue<string>() == userName);
                    if (row != null)
                    {
                        student.StudentName = row.Cell(1).GetValue<string>();
                        student.StudentEmail = row.Cell(2).GetValue<string>();
                        student.StudentPhone = row.Cell(3).GetValue<string>();
                    }
                }
            }
            return student;
        }







        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
