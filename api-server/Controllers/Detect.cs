using Microsoft.AspNetCore.Mvc;
using System;

namespace APIServer.Controllers
{
    [Route("api/detect")]
    [ApiController]
    public class Detect : ControllerBase
    {
        // 用于检测web服务是否正常运行
        [HttpGet("index")]
        public string Index()
        {
            Console.WriteLine("detect");
            return "ok";
        }
    }
}
