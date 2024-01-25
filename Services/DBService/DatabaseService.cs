using ExcelWorkbookPivotTable.ProtectData;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Reflection;

namespace ExcelWorkbookPivotTable.Services.DBService
{
    public interface IDatabaseService
    {
        DataTable GetDataForPivot(Models.RequestModel.UserRequest request);

        DataTable GenerateRandomDataTable(string startDate, string endDate, int rowCount);
        string EncryptData(string plainText);
    }

    public class DatabaseService : IDatabaseService
    {

        private readonly ILogger<DatabaseService> _logger;
        private IConfiguration _configuration { get; set; }

        private readonly ISecureInfo _secureInfo;

        private static string dbConnectionSecret;
        public DatabaseService(ILogger<DatabaseService> logger, IConfiguration configuration, ISecureInfo secureInfo)
        {
            _logger = logger;
            _configuration = configuration;
            _secureInfo = secureInfo;
            dbConnectionSecret = _configuration["ConnectionStrings:DbConnection"];
        }
        #region Constants
        //List of City names
        List<string> cities = new List<string>
            {
                "New York", "London", "Paris", "Tokyo", "Sydney", "Berlin", "Mumbai", "Beijing", "Rio de Janeiro", "Cairo",
                "Moscow", "Los Angeles", "Toronto", "Dubai", "Rome", "Seoul", "Cape Town", "Buenos Aires", "Istanbul", "Mexico City"
            };
        // List of human names
        List<string> names = new List<string>
            {
                "John", "Jane", "Michael", "Emily", "David", "Sophia", "Daniel", "Olivia", "Matthew", "Ava",
                "Christopher", "Emma", "Andrew", "Isabella", "Ethan", "Mia", "William", "Abigail", "James", "Ella"
                // Add more names as needed
            };
        //List of Originator names
        List<string> originators = new List<string>
            {
                "Alice", "Bob", "Charlie", "Diana", "Edward", "Fiona", "George", "Helen", "Ivan", "Julia",
                "Kevin", "Linda", "Mark", "Nina", "Oscar", "Pamela", "Quincy", "Rachel", "Samuel", "Tina"
                // Add more originator names as needed
            };
        #endregion


        #region To get Dynamic data from Database using SP, based on your data
        //to get the data for the Pivot sheet and Dashboard sheet from the database using SP
        public DataTable GetDataForPivot(Models.RequestModel.UserRequest request)
        {
            DataTable dataTable = new DataTable();

            try
            {
                _logger.LogInformation($"Referral Reporting Details retrieval has started...");
                // encryot firsttt 
                //var sample = EncryptData(dbConnectionSecret);

                var dbConnectionSecretKey = _secureInfo.DecryptData(dbConnectionSecret);
                //string connectionString = _configuration.GetConnectionString("DbConnection");

                using (SqlConnection connection = new SqlConnection(dbConnectionSecretKey))
                {
                    connection.Open();

                    string cmdStr = "dbo.GetReferralDataByDateRange"; // Enter SP name here
                    using (SqlCommand command = new SqlCommand(cmdStr, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@StartDate", DateTime.Parse(request.startDate));
                        command.Parameters.AddWithValue("@EndDate", DateTime.Parse(request.endDate));

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }

                    _logger.LogInformation($"Database connection is closed successfully...");

                    if (dataTable.Rows.Count > 0)
                    {
                        _logger.LogInformation($"Referral Reporting Details retrieved successfully.");
                    }
                    else
                    {
                        _logger.LogInformation($"No Referral Reporting Details.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Referral Reporting Details retrieval has failed.");
            }

            return dataTable;
        }
        #endregion


        //to get the data for Excel using manual data creation
        #region Generate Data Using Hardcoding data
        public DataTable GenerateRandomDataTable(string startDate, string endDate, int rowCount)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Date", typeof(string));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Source", typeof(string));
            dataTable.Columns.Add("Originator", typeof(string));
            dataTable.Columns.Add("Comments", typeof(string));

            Random random = new Random();

            // Generate a list of random city names
            List<string> randomCities = GenerateRandomCityList(rowCount); // Assuming a maximum of 10 cities
            List<string> randomNames = GenerateRandomNameList(rowCount);
            List<string> randomOriginators = GenerateRandomOriginatorList(rowCount);


            DateTime startDateTime = DateTime.ParseExact(startDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            DateTime endDateTime = DateTime.ParseExact(endDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            while (startDateTime <= endDateTime)
            {
                DataRow row = dataTable.NewRow();
                row["Date"] = startDateTime.ToString("MMMM", CultureInfo.InvariantCulture);
                row["Name"] = randomNames[random.Next(randomNames.Count)];
                row["Source"] = randomCities[random.Next(randomCities.Count)];
                row["Originator"] = randomOriginators[random.Next(randomOriginators.Count)];
                row["Comments"] = GetRandomString(8);

                dataTable.Rows.Add(row);
                startDateTime = startDateTime.AddDays(1);
            }

            return dataTable;
        }


        private List<string> GenerateRandomCityList(int maxCityCount)
        {
            // Shuffle the list to get random cities
            cities = cities.OrderBy(x => Guid.NewGuid()).ToList();
            // Take a random number of cities, up to the specified maximum count
            int cityCount = new Random().Next(1, maxCityCount + 1);
            return cities.Take(cityCount).ToList();
        }

        private List<string> GenerateRandomNameList(int maxNameCount)
        {

            names = names.OrderBy(x => Guid.NewGuid()).ToList();
            int nameCount = new Random().Next(1, maxNameCount + 1);
            return names.Take(nameCount).ToList();
        }

        private List<string> GenerateRandomOriginatorList(int maxOriginatorCount)
        {

            originators = originators.OrderBy(x => Guid.NewGuid()).ToList();
            int originatorCount = new Random().Next(1, maxOriginatorCount + 1);
            return originators.Take(originatorCount).ToList();
        }

        private string GetRandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            Random random = new Random();
            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }
        #endregion






        #region To Encrypt the appsettings data 
        public string EncryptData(string plainText)
        {
            try
            {
                if (string.IsNullOrEmpty(plainText))
                    return string.Empty;

                var encryptedValue = _secureInfo.EncryptData(plainText);
                return encryptedValue;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Data encryption has failed.");
                throw ex;
            }
        }

        // encrypt data and paste in app settings
        //public string EncryptData(string plainText)
        //{
        //    try
        //    {
        //        if (string.IsNullOrEmpty(plainText))
        //            return string.Empty;

        //        string encryptedString = _secureInfo.EncryptData(plainText);

        //        // Read the existing appsettings.json content into a Dictionary
        //        var configPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
        //        var json = File.ReadAllText(configPath);
        //        var appSettings = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

        //        // Update the DbConnection property
        //        appSettings["ConnectionStrings"] = new Dictionary<string, object>
        //            {
        //                { "DbConnection", encryptedString }
        //            };

        //        // Serialize the updated Dictionary back to JSON
        //        var updatedJson = JsonConvert.SerializeObject(appSettings, Newtonsoft.Json.Formatting.Indented);

        //        // Write the updated JSON back to appsettings.json
        //        File.WriteAllText(configPath, updatedJson);

        //        return encryptedString;
        //    }
        //    catch (Exception ex)
        //    {
        //        // Handle exceptions
        //        throw ex;
        //    }
        //}

        #endregion
    }
}
