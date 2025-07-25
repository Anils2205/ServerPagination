using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ServerPagination.Models;
using ServerPagination.Models.Comman;
using ServerPagination.Services.Abstraction;
using System;
namespace ServerPagination.API.Controllers
{
    [Route("api/v1")]
    [ApiController]
    public class HomeApiController : Controller
    {
        private readonly IHomeHelper _homeHelper;
        public HomeApiController(IHomeHelper homeHelper)
        {
            _homeHelper = homeHelper;
        }
        [Route("userList")]
        [HttpPost]
        public IActionResult UserList(SetPagination setPagination)
        {
            Response response = new();
            try
            {
                var userdata = _homeHelper.UserList(setPagination);
                if (userdata != null)
                {
                    response.code = StatusCodes.Status200OK;
                    response.status = true;
                    response.message = "Success";
                    response.data = userdata;
                    return Ok(response);
                }
                response.code = StatusCodes.Status400BadRequest;
                response.status = false;
                response.message = "Object is null.";
                return BadRequest(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
        [Route("addUser")]
        [HttpPost]
        public IActionResult AddUser(UserModel user)
        {
            Response response = new();
            try
            {
                _homeHelper.AddUser(user);
                response.code = StatusCodes.Status200OK;
                response.status = true;
                response.message = "Success";
                return Ok(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
        [Route("getUser{UserId}")]
        [HttpGet]
        public IActionResult GetUser(int UserId)
        {
            Response response = new();
            try
            {
                var userdata = _homeHelper.GetUser(UserId);
                if (userdata != null)
                {
                    response.code = StatusCodes.Status200OK;
                    response.status = true;
                    response.message = "Success";
                    response.data = userdata;
                    return Ok(response);
                }
                response.code = StatusCodes.Status400BadRequest;
                response.status = false;
                response.message = "Object is null.";
                return BadRequest(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
        [Route("editUser")]
        [HttpPut]
        public IActionResult EditUser(EditUserModel user)
        {
            Response response = new();
            try
            {
                _homeHelper.EditUser(user);
                response.code = StatusCodes.Status200OK;
                response.status = true;
                response.message = "Success";
                return Ok(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
        [Route("activeManage{UserId}")]
        [HttpGet]
        public IActionResult ActiveManage(int UserId)
        {
            Response response = new();
            try
            {

                _homeHelper.ActiveManage(UserId);
                response.code = StatusCodes.Status200OK;
                response.status = true;
                response.message = "Success";
                return Ok(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
        [Route("deleteUser{UserId}")]
        [HttpDelete]
        public IActionResult DeleteUser(int UserId)
        {
            Response response = new();
            try
            {
                _homeHelper.DeleteUser(UserId);
                response.code = StatusCodes.Status200OK;
                response.status = true;
                response.message = "Success";
                return Ok(response);
            }
            catch (Exception e)
            {
                response.code = StatusCodes.Status500InternalServerError;
                response.status = false;
                response.message = "Something went wrong." + e.Message;
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
    }
}