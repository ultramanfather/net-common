using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using System.IO;

namespace APIServer
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseStaticFiles();

            // ��ʼ����̬�ļ�Ŀ¼
            string rootStaticDir = Configuration.GetValue<string>("StaticFileDir").Replace("<exec>", System.AppDomain.CurrentDomain.BaseDirectory);
            // �����ľ�̬�ļ�Ŀ¼
            string exportStaticDir = Path.Combine(rootStaticDir, Configuration.GetValue<string>("ExportDirName"));
            if (!Directory.Exists(exportStaticDir))
            {
                Directory.CreateDirectory(exportStaticDir);
            }

            // ��̬�ļ�����
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = new PhysicalFileProvider(rootStaticDir),
                RequestPath = Configuration.GetValue<string>("StaticFileRequestPath")
            });

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
