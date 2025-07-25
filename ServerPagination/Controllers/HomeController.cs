using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using ServerPagination.Models;
using ServerPagination.Models.Comman;
using ServerPagination.Services;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace ServerPagination.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<HomeController> _logger;
        private readonly IManageService<SetPagination, PaginationModel> _userListManageService;
        private readonly IManageService<UserModel, string> _addUserManageService;
        private readonly IManageService<(string, int), string> _deleteUserManageService;
        private readonly IManageService<(string, int), string> _activeManageService;
        private readonly IManageService<(string, int), EditUserModel> _getUserManageService;
        private readonly IManageService<EditUserModel, string> _editUserManageService;
        public HomeController
            (
                IManageService<SetPagination, PaginationModel> userListManageService,
                IManageService<UserModel, string> addUserManageService,
                IManageService<(string, int), string> deleteUserManageService,
                IManageService<(string, int), string> activeManageService,
                IManageService<(string, int), EditUserModel> getUserManageService,
                IManageService<EditUserModel, string> editUserManageService,
                IWebHostEnvironment env, ILogger<HomeController> logger
            )
        {
            _userListManageService = userListManageService;
            _addUserManageService = addUserManageService;
            _deleteUserManageService = deleteUserManageService;
            _activeManageService = activeManageService;
            _getUserManageService = getUserManageService;
            _editUserManageService = editUserManageService;
            _env = env;
            _logger = logger;
    }
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Profile()
        {
            return View();
        }
        [HttpGet]
        public IActionResult Users()
        {
            return View();
        }
        [HttpGet]
        public async Task<IActionResult> LoadUserList(int pageNumber, string searchQuery)
        {
            try
            {
                SetPagination setPagination = new SetPagination();
                setPagination.PageNumber = pageNumber;
                if (searchQuery != null)
                {
                    setPagination.SearchQuery = searchQuery;
                }
                var response = await _userListManageService.PostAsync(UrlConstants.UserListUrl, setPagination);
                return PartialView("_UserList", response);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        [HttpGet]
        public IActionResult AddUser()
        {
            return PartialView("_AddUser");
        }
        [HttpPost]
        public IActionResult AddUser(UserModel user)
        {
            try
            {
                TryValidateModel(user);
                if (ModelState.IsValid)
                {
                    var response = _addUserManageService.PostAsync(UrlConstants.AddUserUrl, user);
                    if (response == null)
                    {
                        return BadRequest();
                    }
                    return Ok(true);
                }
                return View(user);
                
            }
            catch (Exception ex)
            {
                return View(ex.Message);
            }
        }
        [HttpGet]
        public IActionResult GetUser(int UserId)
        {
            var response = _getUserManageService.GetAsync(UrlConstants.GetUserManageUrl, UserId);
            if (response == null)
            {
                return BadRequest();
            }
            return PartialView("_EditUser", response.Result);
        }
        [HttpPut]
        public IActionResult EditUser(EditUserModel user)
        {
            TryValidateModel(user);
            if (ModelState.IsValid)
            {
                var response = _editUserManageService.PutAsync(UrlConstants.EditUserUrl, user);
                if (response == null)
                {
                    return BadRequest();
                }
                return Ok(true);
            }
            return PartialView("_EditUser", user);
        }
        [HttpGet]
        public IActionResult ActiveManage(int UserId)
        {
            var response = _activeManageService.GetAsync(UrlConstants.ActiveManageUrl, UserId);
            if (response == null)
            {
                return BadRequest();
            }
            return Ok(true);
        }
        [HttpDelete]
        public IActionResult DeleteUser(int UserId)
        {
            var response = _deleteUserManageService.DeleteAsync(UrlConstants.DeleteUserUrl, UserId);
            if (response == null)
            {
                return BadRequest();
            }
            return Ok(true);
        }
        [HttpGet]
        public IActionResult ViewUser(int UserId)
        {
            var response = _getUserManageService.GetAsync(UrlConstants.GetUserManageUrl, UserId);
            if (response == null)
            {
                return BadRequest();
            }
            return PartialView("_ViewUser", response.Result);
        }
        [HttpGet]
        public IActionResult CompanyDetails()
        {
            return View();
        }
        [HttpGet]
        public IActionResult DownloadExcel()
        {
            try
            {
                // Set the license context to non-commercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Log the environment web root path
                _logger.LogInformation($"Web root path: {_env.WebRootPath}");

                // Use a secure temporary directory for file generation
                var tempDirectory = Path.Combine(_env.WebRootPath, "reports");
                if (!Directory.Exists(tempDirectory))
                {
                    _logger.LogInformation($"Creating directory: {tempDirectory}");
                    Directory.CreateDirectory(tempDirectory);
                }

                var filePath = Path.Combine(tempDirectory, "report.xlsx");
                _logger.LogInformation($"File path: {filePath}");

                // Generate the Excel file
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Report");

                    worksheet.Cells[1, 1].Value = "ID";
                    worksheet.Cells[1, 2].Value = "Name";
                    worksheet.Cells[1, 3].Value = "Age";

                    worksheet.Cells[2, 1].Value = 1;
                    worksheet.Cells[2, 2].Value = "John Doe";
                    worksheet.Cells[2, 3].Value = 30;

                    worksheet.Cells[3, 1].Value = 2;
                    worksheet.Cells[3, 2].Value = "Jane Smith";
                    worksheet.Cells[3, 3].Value = 25;

                    // Save the file to the specified location
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                    _logger.LogInformation($"Excel file saved: {filePath}");
                }

                // Send the file to the client for download
                var memory = new MemoryStream();
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    stream.CopyTo(memory);
                }
                memory.Position = 0;

                // Clean up temporary file
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                    _logger.LogInformation($"Temporary file deleted: {filePath}");
                }

                return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating or downloading Excel file.");
                return StatusCode(500, "Internal server error");
            }
        }
        [HttpGet]
        public IActionResult DownloadPdf()
        {
            try
            {
                var userData = _getUserManageService.GetAsync(UrlConstants.GetUserManageUrl, 10);
                return File(DownloadInvoice(userData.Result), "application/pdf", "Invoice.pdf");
            }   
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating or downloading Excel file.");
                return StatusCode(500, "Internal server error");
            }
        }
        public byte[] DownloadInvoice(EditUserModel invoiceDetails)
        {
            string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(Convert.ToInt32(invoiceDetails.CreatedDate?.ToString("MM")));
            var memoryStream = new MemoryStream();

            // Marge in centimeter, then I convert with .ToDpi()
            float margeLeft = 1.5f;
            float margeRight = 1.5f;
            float margeTop = 0.5f;
            float margeBottom = 1.0f;

            Document pdf = new(
                                    PageSize.A4,
                                    margeLeft.ToDpi(),
                                    margeRight.ToDpi(),
                                    margeTop.ToDpi(),
                                    margeBottom.ToDpi()
                                   );

            pdf.AddTitle("InvoiceDetails");
            pdf.AddAuthor("RentEasly");
            pdf.AddCreationDate();
            pdf.AddKeywords("RentEasly");
            pdf.AddSubject("InvoiceDetails");

            //string backgroundImageFilePath = "https://bonjonsystem.s3.us-east-2.amazonaws.com/images/logo.png";
            //Image imgBackground = Image.GetInstance(backgroundImageFilePath);
            //imgBackground.ScaleToFit(500, 500);
            //imgBackground.Alignment = Image.UNDERLYING;
            //imgBackground.SetAbsolutePosition(0, 0);

            PdfWriter writer = PdfWriter.GetInstance(pdf, memoryStream);

            pdf.Open();
            PdfContentByte under = writer.DirectContentUnder;
            //under.AddImage(imgBackground);

            string logoImageURL = "https://bonjonsystem.s3.us-east-2.amazonaws.com/images/logo.png";
            Image imglogo = Image.GetInstance(logoImageURL);
            imglogo.ScaleToFit(200f, 200f);
            imglogo.SpacingBefore = 10f;
            imglogo.SpacingAfter = 1f;
            imglogo.Alignment = Element.ALIGN_LEFT;

            PdfPTable maindatatable = new PdfPTable(6);
            float[] maindatatableheaderwidths = { 20, 20, 15, 15, 15, 15 };
            maindatatable.WidthPercentage = 100;
            maindatatable.SetWidths(maindatatableheaderwidths);

            PdfPTable footerdatatable = new PdfPTable(6);
            float[] footerdatatableheaderwidths = { 20, 20, 15, 15, 15, 15 };
            footerdatatable.WidthPercentage = 100;
            footerdatatable.SetWidths(footerdatatableheaderwidths);

            PdfPCell apartmentImageCell = new PdfPCell();
            apartmentImageCell.Colspan = 1;
            apartmentImageCell.HorizontalAlignment = Element.ALIGN_LEFT;
            apartmentImageCell.VerticalAlignment = Element.ALIGN_TOP;
            apartmentImageCell.Image = imglogo;
            apartmentImageCell.PaddingBottom = 0;
            apartmentImageCell.PaddingTop = 0;
            apartmentImageCell.PaddingRight = 15;
            apartmentImageCell.BorderWidthTop = 0;
            apartmentImageCell.BorderWidthRight = 0;
            apartmentImageCell.BorderWidthLeft = 0;
            apartmentImageCell.BorderWidthBottom = 0;
            apartmentImageCell.FixedHeight = 60;
            maindatatable.AddCell(apartmentImageCell);

            PdfPCell empty = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 14, 1, BaseColor.WHITE))) { BorderColor = BaseColor.WHITE, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 10, PaddingTop = 0, PaddingBottom = 0 };
            empty.Colspan = 2;
            maindatatable.AddCell(empty);
            PdfPCell rightheader = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.WHITE, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 0, PaddingBottom = 0 };
            rightheader.Colspan = 3;
            maindatatable.AddCell(rightheader);
            PdfPCell title = new PdfPCell(new Phrase("www.bonjansystem.com", new Font(Font.FontFamily.TIMES_ROMAN, 14, 1, BaseColor.BLUE))) { BorderColor = BaseColor.WHITE, BorderWidthBottom = 1, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 0, PaddingBottom = 10, PaddingLeft = 5 };
            title.Colspan = 6;

            maindatatable.AddCell(title);
            Font normalFont = new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, new BaseColor(System.Drawing.Color.Black));
            Font boldfont = new Font(Font.FontFamily.TIMES_ROMAN, 16, 1, new BaseColor(System.Drawing.Color.Black));

            PdfPCell c1 = new PdfPCell(new Phrase("Bill To", boldfont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            c1.Colspan = 3;
            maindatatable.AddCell(c1);
            PdfPCell invoiceNo = new PdfPCell(new Phrase("Invoice - #" + 05154, new Font(Font.FontFamily.TIMES_ROMAN, 14, 1, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.WHITE, BorderWidthBottom = 1, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            invoiceNo.Colspan = 3;
            maindatatable.AddCell(invoiceNo);

            PdfPCell c2 = new PdfPCell(new Phrase("ID: " + 3 , normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            c2.Colspan = 3;
            maindatatable.AddCell(c2);

            PdfPCell date = new PdfPCell(new Phrase("Date: " + DateTime.Now.ToString("MMM dd,yyyy"), normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            date.Colspan = 3;
            maindatatable.AddCell(date);

            PdfPCell c3 = new PdfPCell(new Phrase("Name: " + invoiceDetails.FirstName + invoiceDetails.LastName, normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            c3.Colspan = 3;
            maindatatable.AddCell(c3);


            PdfPCell duedate = new PdfPCell(new Phrase("Start Date: " + invoiceDetails.CreatedDate?.ToString("MMM dd,yyyy"), normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            duedate.Colspan = 3;
            maindatatable.AddCell(duedate);

            PdfPCell c4 = new PdfPCell(new Phrase( "Phone: " + invoiceDetails.MobileNo , normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            c4.Colspan = 3;
            maindatatable.AddCell(c4);

            PdfPCell leaseenddate = new PdfPCell(new Phrase("End Date: " + invoiceDetails.UpdatedDate?.ToString("MMM dd,yyyy"), normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            leaseenddate.Colspan = 3;
            maindatatable.AddCell(leaseenddate);

            PdfPCell c6 = new PdfPCell(new Phrase("Email:" + invoiceDetails.EmailAddress, normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 40 };
            c6.Colspan = 6;
            maindatatable.AddCell(c6);


            //***** Table 1 *****
            PdfPCell srno = new PdfPCell(new Phrase("SR No.", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            srno.Colspan = 1;
            maindatatable.AddCell(srno);

            PdfPCell title1 = new PdfPCell(new Phrase("Title", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            title1.Colspan = 2;
            maindatatable.AddCell(title1);

            PdfPCell description = new PdfPCell(new Phrase("Description", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            description.Colspan = 2;
            maindatatable.AddCell(description);

            PdfPCell amount = new PdfPCell(new Phrase("Amount", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            amount.Colspan = 1;
            maindatatable.AddCell(amount);

            PdfPCell srnovalue1 = new PdfPCell(new Phrase( "1" , normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            srnovalue1.Colspan = 1;
            maindatatable.AddCell(srnovalue1);

            PdfPCell titlevalue1 = new PdfPCell(new Phrase("Title1", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            titlevalue1.Colspan = 2;
            maindatatable.AddCell(titlevalue1);

            PdfPCell descriptionvalue1 = new PdfPCell(new Phrase("Description... 1" + monthName, normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            descriptionvalue1.Colspan = 2;
            maindatatable.AddCell(descriptionvalue1);

            PdfPCell amountvalue1 = new PdfPCell(new Phrase("3", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            amountvalue1.Colspan = 1;
            maindatatable.AddCell(amountvalue1);

            PdfPCell srnovalue2 = new PdfPCell(new Phrase("2", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            srnovalue2.Colspan = 1;
            maindatatable.AddCell(srnovalue2);

            PdfPCell titlevalue2 = new PdfPCell(new Phrase("Title2", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            titlevalue2.Colspan = 2;
            maindatatable.AddCell(titlevalue2);

            PdfPCell descriptionvalue2 = new PdfPCell(new Phrase("Description... 2" + monthName, normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            descriptionvalue2.Colspan = 2;
            maindatatable.AddCell(descriptionvalue2);

            PdfPCell amountvalue2 = new PdfPCell(new Phrase("2400", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            amountvalue2.Colspan = 1;
            maindatatable.AddCell(amountvalue2);

            PdfPCell empty4 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_CENTER, PaddingLeft = 5, PaddingTop = 80, PaddingBottom = 80 };
            empty4.Colspan = 6;
            maindatatable.AddCell(empty4);

            //***** Table 2 *****

            PdfPCell venderId = new PdfPCell(new Phrase("VenderId", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            venderId.Colspan = 1;
            maindatatable.AddCell(venderId);

            PdfPCell vanderName  = new PdfPCell(new Phrase("Vander Name", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            vanderName.Colspan = 2;
            maindatatable.AddCell(vanderName);

            PdfPCell phoneNo = new PdfPCell(new Phrase("PhoneNo", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            phoneNo.Colspan = 2;
            maindatatable.AddCell(phoneNo);

            PdfPCell venderTicketSoldCount = new PdfPCell(new Phrase("Ticket Sold Count", new Font(Font.FontFamily.TIMES_ROMAN, 14, 0, BaseColor.DARK_GRAY))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 1, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingRight = 20, PaddingTop = 15, PaddingBottom = 15 };
            venderTicketSoldCount.Colspan = 1;
            maindatatable.AddCell(venderTicketSoldCount);

            PdfPCell venderIdvalue1 = new PdfPCell(new Phrase("54", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            venderIdvalue1.Colspan = 1;
            maindatatable.AddCell(venderIdvalue1);

            PdfPCell vanderNamevalue1 = new PdfPCell(new Phrase("Disha", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            vanderNamevalue1.Colspan = 2;
            maindatatable.AddCell(vanderNamevalue1);

            PdfPCell phoneNovalue1 = new PdfPCell(new Phrase("+91 542526505498", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            phoneNovalue1.Colspan = 2;
            maindatatable.AddCell(phoneNovalue1);

            PdfPCell venderTicketSoldCountvalue1 = new PdfPCell(new Phrase("3", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            venderTicketSoldCountvalue1.Colspan = 1;
            maindatatable.AddCell(venderTicketSoldCountvalue1);

            PdfPCell venderIdvalue2 = new PdfPCell(new Phrase("65", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            venderIdvalue2.Colspan = 1;
            maindatatable.AddCell(venderIdvalue2);

            PdfPCell vanderNamevalue2 = new PdfPCell(new Phrase("Kim", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            vanderNamevalue2.Colspan = 2;
            maindatatable.AddCell(vanderNamevalue2);

            PdfPCell phoneNovalue2 = new PdfPCell(new Phrase("+91 5142942446", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            phoneNovalue2.Colspan = 2;
            maindatatable.AddCell(phoneNovalue2);

            PdfPCell venderTicketSoldCountvalue2 = new PdfPCell(new Phrase("2", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_RIGHT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 10 };
            venderTicketSoldCountvalue2.Colspan = 1;
            maindatatable.AddCell(venderTicketSoldCountvalue2);

            PdfPCell empty5 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_CENTER, PaddingLeft = 5, PaddingTop = 80, PaddingBottom = 80};
            empty5.Colspan = 6;
            maindatatable.AddCell(empty5);

            PdfPCell termsvalue = new PdfPCell(new Phrase("If you have any questions about this invoice, please contact", normalFont)) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_CENTER, PaddingLeft = 5, PaddingTop = 40, PaddingBottom = 10 };
            termsvalue.Colspan = 6;
            maindatatable.AddCell(termsvalue);

            PdfPCell supoort = new PdfPCell(new Phrase("support@bonjansystem.com", boldfont)) { BorderColor = BaseColor.BLUE, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_CENTER, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 20 };
            supoort.Colspan = 6;
            maindatatable.AddCell(supoort);

            //string apartmentImagePath = "https://green-bush-021956c00.azurestaticapps.net/assets/img/Apartment rent-bro-01.png";
            //Image imgApartment = Image.GetInstance(apartmentImagePath);
            //imgApartment.ScaleAbsolute(10f, 10f);
            //imgApartment.SpacingBefore = 0f;
            //imgApartment.SpacingAfter = 0f;
            //imgApartment.Alignment = Element.ALIGN_TOP;

            //PdfPCell cellImage = new PdfPCell(imgApartment, true);
            //cellImage.Colspan = 3;
            //cellImage.Rowspan = 6;
            //cellImage.HorizontalAlignment = Element.ALIGN_LEFT;
            //cellImage.VerticalAlignment = Element.ALIGN_TOP;
            //cellImage.PaddingRight = 20;
            //cellImage.BorderWidthTop = 0;
            //cellImage.BorderWidthRight = 0;
            //cellImage.BorderWidthLeft = 0;
            //cellImage.BorderWidthBottom = 0;
            //cellImage.FixedHeight = 50;



            //PdfPCell PropertyName = new PdfPCell(new Phrase("Property Name:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(PropertyName);

            ////PdfPCell PropertyNameValue = new PdfPCell(new Phrase(propertyDetails.BuildingDetails.PropertyName, new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////dtContent.AddCell(PropertyNameValue);

            //PdfPCell PropertyAddress = new PdfPCell(new Phrase("Property Address:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(PropertyAddress);

            ////PdfPCell PropertyAddressValue = new PdfPCell(new Phrase(propertyDetails.BuildingDetails.AddrLineOne.ToString() + "\n" + propertyDetails.BuildingDetails.AddrLineTwo.ToString(), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////dtContent.AddCell(PropertyAddressValue);

            //PdfPCell city = new PdfPCell(new Phrase("City:  Lorem Ipsum", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(city);

            //PdfPCell zipcode = new PdfPCell(new Phrase("zip code:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(zipcode);


            //PdfPCell LeasePeriodStartDate = new PdfPCell(new Phrase("Lease Terms:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(LeasePeriodStartDate);

            ////PdfPCell LeasePeriodStartDateValue = new PdfPCell(new Phrase("Begin Date: " + propertyDetails.LeaseDetails.StartDate.ToString("dd/MM/yyyy"), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////dtContent.AddCell(LeasePeriodStartDateValue);

            //PdfPCell empty2 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //dtContent.AddCell(empty2);

            ////PdfPCell LeasePeriodEndDateValue = new PdfPCell(new Phrase("End Date: " + propertyDetails.LeaseDetails.EndDate.ToString("dd/MM/yyyy"), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////dtContent.AddCell(LeasePeriodEndDateValue);



            //PdfPCell ContentCell = new PdfPCell(dtContent) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 35, PaddingBottom = 20, };
            //ContentCell.Colspan = 3;

            //maindatatable.AddCell(ContentCell);
            //maindatatable.AddCell(cellImage);


            //PdfPCell SecurityDeposit = new PdfPCell(new Phrase("Security Deposit:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //footerdatatable.AddCell(SecurityDeposit);

            ////PdfPCell SecurityDepositValue = new PdfPCell(new Phrase("Rs " + propertyDetails.LeaseDetails.SecurityDeposit.ToString("00.00", AppConstants.DefaultCulture), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////footerdatatable.AddCell(SecurityDepositValue);

            //PdfPCell Status = new PdfPCell(new Phrase("Status:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            //footerdatatable.AddCell(Status);

            ////PdfPCell StatusValue = new PdfPCell(new Phrase(propertyDetails.LeaseDetails.IsSecurityPaid ? "Paid" : "Pending", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////footerdatatable.AddCell(StatusValue);

            ////PdfPCell SecurityDepositDueDate = new PdfPCell(new Phrase("Due Date: " + (propertyDetails.LeaseDetails.IsSecurityPaid ? "" : $"Due Date: {propertyDetails.LeaseDetails.SecDepDueDate:dd/MM/yyyy}").Replace("00:00:00", ""), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////footerdatatable.AddCell(SecurityDepositDueDate);

            ////PdfPCell RentDueDateValue = new PdfPCell(new Phrase(propertyDetails.LeaseDetails.SecDepDueDate.ToString(), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 10 };
            ////footerdatatable.AddCell(RentDueDateValue);

            ////PdfPCell MonthlyRent = new PdfPCell(new Phrase("Monthly Rent:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 20 };
            ////footerdatatable.AddCell(MonthlyRent);

            ////PdfPCell MonthlyRentValue = new PdfPCell(new Phrase("Rs " + propertyDetails.LeaseDetails.MonthlyRent.ToString("00.00", AppConstants.DefaultCulture), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 20 };
            ////footerdatatable.AddCell(MonthlyRentValue);

            ////PdfPCell DueBy = new PdfPCell(new Phrase("Due By:" + propertyDetails.LeaseDetails.RentDueDate?.Ordinalize(), new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 20 };
            ////footerdatatable.AddCell(DueBy);

            ////PdfPCell empty6 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 0, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 0, PaddingBottom = 20 };
            ////empty6.Colspan = 3;
            ////footerdatatable.AddCell(empty6);

            ////PdfPCell Note = new PdfPCell(new Phrase("Note:", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 0 };
            ////Note.Colspan = 6;
            ////datatable1.AddCell(Note);


            //PdfPCell NoteValue = new PdfPCell(new Phrase("Note: This is an electronic copy of your rental summary and does not require any authorised signatory.For any discrepancies,Please Contact your Property owner via call or email.", new Font(Font.FontFamily.TIMES_ROMAN, 12, 0, BaseColor.BLACK))) { BorderColor = BaseColor.BLACK, BorderWidthBottom = 0, BorderWidthTop = 1, BorderWidthRight = 0, BorderWidthLeft = 0, HorizontalAlignment = Element.ALIGN_LEFT, PaddingLeft = 5, PaddingTop = 10, PaddingBottom = 0 };
            //NoteValue.Colspan = 6;
            //footerdatatable.AddCell(NoteValue);

            pdf.Add(maindatatable);
            //pdf.Add(footerdatatable);
            pdf.Close();
            return memoryStream.ToArray();
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
    public static class Extensions
    {
        public static float ToDpi(this float centimeter)
        {
            var inch = centimeter / 2.54;
            return (float)(inch * 72);
        }
    }
}