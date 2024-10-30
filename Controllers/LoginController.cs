using Class;
using Microsoft.AspNetCore.Mvc;
using Model;
using Microsoft.AspNetCore.Antiforgery;
using System.Web;
using System.Data;
using Microsoft.AspNetCore.Http;
using System.Diagnostics;

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginController : ControllerBase
    {
        private readonly IAntiforgery _antiforgery;
        private readonly IHttpContextAccessor _httpContextAccessor;
        public LoginController(IAntiforgery antiforgery, IHttpContextAccessor httpContextAccessor)
        {
            _antiforgery = antiforgery;
            _httpContextAccessor = httpContextAccessor;
        }
        //[IgnoreAntiforgeryToken]  
        //[HttpGet("convert_to_pdf", Name = "convert_to_pdf")]
        //public IActionResult convert_to_pdf()
        //{
        //    string msg = "";
        //    string libreOfficePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "LibreOffice", "program", "soffice.exe");
        //    ProcessStartInfo startInfo = new ProcessStartInfo();
        //    startInfo.FileName = libreOfficePath;// "soffice.exe";
        //    startInfo.Arguments = @"--headless --convert-to pdf C:\Users\2bLove\source\repos\dotnet-epha-api-8\tools\LibreOffice\program\file.xlsx";
        //    startInfo.RedirectStandardOutput = true;
        //    startInfo.RedirectStandardError = true;
        //    startInfo.UseShellExecute = false;
        //    startInfo.CreateNoWindow = true;

        //    using (Process process = Process.Start(startInfo))
        //    {
        //        bool exited = process.WaitForExit(30000); // รอให้กระบวนการทำงานเสร็จภายใน 30 วินาที (30000 มิลลิวินาที)

        //        if (!exited)
        //        {
        //            // ถ้ากระบวนการไม่เสร็จภายในเวลาที่กำหนด ให้บังคับปิด
        //            process.Kill();
        //            throw new Exception("Process timed out and was killed.");
        //        }
        //        string result = process.StandardOutput.ReadToEnd();
        //        string error = process.StandardError.ReadToEnd();

        //        if (!string.IsNullOrEmpty(error))
        //        {
        //            throw new Exception("Error: " + error);
        //        }
        //    } 
        //    return Ok(new { msg });
        //}


        // ขั้นตอนที่ 1: สร้าง CSRF Token โดยข้ามการตรวจสอบ CSRF
        [IgnoreAntiforgeryToken]  // ข้ามการตรวจสอบ CSRF
        [HttpGet("GetAntiForgeryToken", Name = "GetAntiForgeryToken")]
        public IActionResult GetAntiForgeryToken()
        {
            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);  // สร้าง CSRF token และเก็บใน Cookie
            return Ok(new { csrfToken = tokens.RequestToken });
        }

        //[IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF เพราะคำขอนี้ใช้สร้าง CSRF Token
        //[HttpGet("ValidateCSRFToken", Name = "ValidateCSRFToken")]
        //public IActionResult ValidateCSRFToken()
        //{
        //    // ตรวจสอบว่ามี Cookie ที่ชื่อ X-CSRF-TOKEN หรือไม่
        //    if (Request.Cookies.TryGetValue("X-CSRF-TOKEN", out var csrfToken))
        //    {
        //        Console.WriteLine("CSRF Token from cookie: " + csrfToken);

        //        // คุณสามารถเพิ่มการตรวจสอบเพิ่มเติมหรือใช้งาน token ได้ที่นี่
        //        if (!string.IsNullOrEmpty(csrfToken))
        //        {
        //            // ถ้าพบค่า CSRF Token ใน cookie
        //            return Ok(new { message = "CSRF token is present", csrfToken });
        //        }
        //        else
        //        {
        //            // ถ้าไม่พบค่าใน CSRF Token
        //            return BadRequest(new { message = "CSRF token is empty" });
        //        }
        //    }
        //    else
        //    {
        //        // ถ้าไม่พบ Cookie ที่ชื่อ X-CSRF-TOKEN
        //        return BadRequest(new { message = "CSRF token not found" });
        //    }
        //}

        //// ขั้นตอนที่ 2: ตรวจสอบ CSRF Token สำหรับคำขอที่ต้องการ
        //[ValidateAntiForgeryToken]
        //[HttpPost("check_authorization_web")]
        //public IActionResult CheckAuthorizationWeb([FromBody] LoginUserModel param)
        //{
        //    try
        //    {
        //        // Log Headers ที่ได้รับ
        //        foreach (var header in Request.Headers)
        //        {
        //            Console.WriteLine($"{header.Key}: {header.Value}");
        //        }

        //        // Log Cookies ที่ได้รับ
        //        foreach (var cookie in Request.Cookies)
        //        {
        //            Console.WriteLine($"{cookie.Key}: {cookie.Value}");
        //        }

        //        if (!ClassLogin.IsAuthorized(param.user_name))
        //        {
        //            return Unauthorized("ผู้ใช้นี้ไม่ได้รับอนุญาต");
        //        }

        //        return Ok(new { message = "ได้รับอนุญาต", user_name = param.user_name });
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error: {ex.Message}");
        //        return StatusCode(500, "Internal server error: " + ex.Message);
        //    }
        //}

        //// ******************* รวมการสร้าง JWT Token และ CSRF Token
        //[IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF
        //[HttpPost("GetAntiForgeryToken")]
        //[ProducesResponseType(typeof(ResultModel<UserResponseViewModel>), StatusCodes.Status200OK)]
        //[Consumes("application/json"), Produces("application/json")]
        //public async Task<IActionResult> GetAntiForgeryToken([FromBody] TokenOneTimeViewModel request)
        //{
        //    try
        //    {
        //        if (!string.IsNullOrEmpty(request.userId))
        //        {
        //            var result = await _authenticationService.GenerateAccessTokenAsync(request);
        //            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);

        //            return Ok(new
        //            {
        //                accessToken = result.accessToken,
        //                csrfToken = tokens.RequestToken
        //            });
        //        }

        //        throw new ProjectServicesException(ExceptionConstant.Code.BadRequest);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new { message = "Error generating tokens", details = ex.Message });
        //    }
        //}


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("check_authorization_page_fix", Name = "check_authorization_page_fix")]
        public string check_authorization_page_fix(PageRoleListModel param)
        {
            ClassLogin cls = new ClassLogin();
            //return cls.check_authorization_page_fix(param);
            string result = cls.check_authorization_page_fix(param);
            return HttpUtility.HtmlEncode(result);

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("check_authorization_page", Name = "check_authorization_page")]
        public string check_authorization_page(PageRoleListModel param)
        {
            ClassLogin cls = new ClassLogin();
            //return cls.authorization_page(param);
            string result = cls.authorization_page(param);

            return HttpUtility.HtmlEncode(result); 
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("check_authorization", Name = "check_authorization")]
        public string check_authorization(LoginUserModel param)
        {
            ClassLogin cls = new ClassLogin();
            //return cls.login(param);
            DataTable dt = new DataTable();
            string result = cls.login(param, ref dt);

            //if (dt != null)
            //{
            //    if (dt.Rows.Count > 0)
            //    {
            //        // ตรวจสอบว่ามี role_type อยู่ใน DataTable หรือไม่
            //        var roleType = dt.Rows[0]["role_type"]?.ToString();
            //        if (!string.IsNullOrEmpty(roleType))
            //        {
            //            // เก็บ role_type ใน Cookie
            //            var cookieOptions = new CookieOptions
            //            {
            //                HttpOnly = true,               // ป้องกันการเข้าถึง cookie จาก JavaScript
            //                Secure = true,                 // ใช้เฉพาะกับ HTTPS เท่านั้น
            //                SameSite = SameSiteMode.Strict, // ป้องกัน CSRF
            //                Expires = DateTime.UtcNow.AddHours(1) // กำหนดอายุของ cookie เป็น 1 ชั่วโมง
            //            };

            //            // role_type ลงใน Cookie  
            //            Response.Cookies.Append("role_type", roleType, cookieOptions);
            //        }
            //    }
            //} 
            return HttpUtility.HtmlEncode(result);
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("register_account", Name = "register_account")]
        public string register_account(RegisterAccountModel param)
        {
            //var roleType = _httpContextAccessor.HttpContext?.Request?.Cookies["role_type"];
            ClassLogin cls = new ClassLogin();
            return cls.register_account(param);

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("update_register_account", Name = "update_register_account")]
        public string update_register_account(RegisterAccountModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.update_register_account(param);

        }

        //[IgnoreAntiforgeryToken]  // ข้ามการตรวจสอบ CSRF
        //[HttpGet("test_export_to_pdf", Name = "test_export_to_pdf")]
        //public string test_export_to_pdf()
        //{
        //    ClassExcel cls = new ClassExcel(); 
        //    string result = cls.test_export_to_pdf();

        //    return HttpUtility.HtmlEncode(result);

        //}
    
    
    }
}
