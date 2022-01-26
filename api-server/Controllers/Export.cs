﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace APIServer.Controllers
{
    [Route("api/export")]
    [ApiController]
    public class Export : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly string staticFileRequestPath;
        private readonly string staticDir;
        public Export(IConfiguration configuration)
        {
            _config = configuration;

            staticDir = _config.GetValue<string>("StaticFileDir").Replace("<exec>", System.AppDomain.CurrentDomain.BaseDirectory);
            staticDir = Path.Combine(staticDir, _config.GetValue<string>("ExportDirName"));
            staticFileRequestPath = _config.GetValue<string>("StaticFileRequestPath") + "/" + _config.GetValue<string>("ExportDirName");
        }

        // HTML 转 Word
        [HttpPost("html-to-word")]
        public string HTML2Word([FromBody] Library.Model.HTML2WordParam param)
        {
            // Word 保存路径
            string filename = Guid.NewGuid().ToString();
            param.Filepath = Path.Combine(staticDir, filename + ".docx");
            param.PDFPath = Path.Combine(staticDir, filename + ".pdf");

            // 生成Word
            OpenXmlConverter.Client.HTML2Word(param);
            string ext;
            if (param.Setting.Convert2PDF == "1")
            {
                ext = ".pdf";
            }
            else
            {
                ext = ".docx";
            }

            return staticFileRequestPath + "/" + filename + ext;
        }
    }
}
