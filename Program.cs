using System.Data.Common;
using System.Data.OleDb;
using Newtonsoft.Json;

const string pathToExcel = @"D:\Downloads\База_Данных_Россия___.xlsx";

const string connectionString = $"""
                                 Provider=Microsoft.ACE.OLEDB.12.0;
                                 Data Source={pathToExcel};
                                 Extended Properties="Excel 12.0 Xml;HDR=YES"
                                 """;

const string sheetMan = "Производители";
const string sheetWin = "Вина";

const string jsonMan = @"D:\Downloads\wine_base_manufacturers.json";
const string jsonWin = @"D:\Downloads\wine_base_wines.json";

await using var dbConnection = new OleDbConnection(connectionString);
await dbConnection.OpenAsync();

await ConvertToJson(dbConnection, sheetMan, jsonMan, GetManufacturers);
await ConvertToJson(dbConnection, sheetWin, jsonWin, GetWines);

return;

async Task ConvertToJson(OleDbConnection connection, string sheetName, string jsonFileName,
    Func<DbDataReader, string> getMethod)
{
    var command = connection.CreateCommand();
    command.CommandText = $"SELECT * FROM [{sheetName}$]";

    await using var reader = await command.ExecuteReaderAsync();

    var json = getMethod(reader);

    await File.WriteAllTextAsync(jsonFileName, json);
}

string GetManufacturers(DbDataReader dataReader)
{
    var data = dataReader
        .Cast<DbDataRecord>()
        .Select(rec => new
        {
            id = rec[0],//Convert.ToInt32(rec[0]),
            guide_year = rec[1],//Convert.ToInt16(rec[1]),
            country = rec[2],
            region = rec[3],
            subregion = rec[4],
            reference_name = rec[5],
            given_name = rec[6],
            area = rec[7],//Convert.ToInt32(rec[7]),
            volume = rec[8],//Convert.ToInt32(rec[8]),
            pictures = new
            {
                id = rec[10],//Convert.ToInt32(rec[10]),
                url = string.Empty // ???
            }
        });
    
    return JsonConvert.SerializeObject(data);
}

string GetWines(DbDataReader dataReader)
{
    var data = dataReader
        .Cast<DbDataRecord>()
        .Select(rec => new
        {
            id = rec[0],//Convert.ToInt32(rec[0]),
            guide_year = rec[1],//Convert.ToInt16(rec[1]),
            //manufacturer_id = Convert.ToInt32(col[2]), // ???
            manufacturer_reference_name = rec[2],
            manufacturer_given_name = rec[3],
            category = rec[4],
            reference_name = rec[5],
            given_name = rec[6],
            age = rec[7],//Convert.ToInt16(rec[7]),
            r = rec[8],//Convert.ToInt16(rec[8]),
            score = rec[9],//Convert.ToInt16(rec[9]),
            reference_page = rec[10],//Convert.ToInt16(rec[10]),
            status = rec[11],
            circulation = rec[12],//Convert.ToInt32(rec[12]),
            pictures = new
            {
                id = rec[14],//Convert.ToInt32(rec[14]),
                url = string.Empty // ???
            },
            type = rec[15],
            grape_sort = rec[16],
            sugar = rec[17],
            resume = rec[18],
            short_description = rec[19],
            recommendations = new
            {
                id = rec[20], //split ???
                text = string.Empty // ???
            },
            alcohol = rec[21],//Convert.ToDecimal(rec[21]),
            production_description = rec[22],
            degustation = rec[23],
            winery_description = rec[24]
        });
    
    return JsonConvert.SerializeObject(data);
}