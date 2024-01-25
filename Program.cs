using ExcelWorkbookPivotTable.ProtectData;
using ExcelWorkbookPivotTable.Services.DBService;
using ExcelWorkbookPivotTable.Services.ExcelService;
using Microsoft.OpenApi.Models;
using Serilog.Events;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddScoped<ISecureInfo, SecureInfo>();
builder.Services.AddScoped<IDatabaseService, DatabaseService>();
builder.Services.AddScoped<IPivotTableLogicService, PivotTableLogicService>();
builder.Services.AddScoped<IWorkbookLogicService, WorkbookLogicService>();

// Add Swagger services
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "ExcelWorkbookPivotTable",
        Version = "v1"
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "ExcelWorkbookPivotTable");
    });
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

// Configure Serilog directly here
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File("log.txt", rollingInterval: RollingInterval.Day)
    .CreateLogger();

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllers();

app.Run();
