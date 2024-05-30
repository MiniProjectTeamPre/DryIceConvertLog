using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static VarInput;

public class ConvertDryice {
    public ConvertDryice() {

    }

    public bool Process(List<DataPackage> data, string path) {
        WriteHeadCsv(path);
        List<DataPackage> dataSort = SortName(data);
        while (dataSort.Count > 0)
        {
            List<DataPackage> dataShort = GetdataShort(ref dataSort);
            List<DataPackage> dataSortDateTime = SortDateTime2(dataShort);
            PrintJson2(path, dataSortDateTime);
        }
        return true;
    }


    private void WriteHeadCsv(string path) {
        try
        {
            using (StreamWriter swOut = new StreamWriter(path, true))
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"ASM,");
                sb.Append($"Customer,");
                sb.Append($"Date,");
                sb.Append($"Time,");
                sb.Append($"Result,");
                sb.Append($"Bat,");
                sb.Append($"Crystal32k,");
                sb.Append($"Crystal8M,");
                sb.Append($"Current,");
                swOut.WriteLine(sb.ToString());
            }
        } catch (Exception ex)
        {
            // Handle any exceptions that may occur during file writing
            // Log or display the error message
        }
    }
    private void PrintJson(string path, List<DataPackage> data) {
        try
        {
            using (StreamWriter swOut = new StreamWriter(path, true))
            {
                foreach (DataPackage dataSup in data)
                {
                    JsonConvert_ json = GetJson(dataSup.Data);
                    StringBuilder sb = new StringBuilder();
                    sb.Append($"{dataSup.Asm},");
                    sb.Append($"{dataSup.Custommer},");
                    sb.Append($"{dataSup.DateTime},");
                    sb.Append($"{dataSup.Result},");

                    foreach (string step in new string[] { "2.1", "3.33", "3.35", "4.4" })
                    {
                        JsonConvert_.ResultString_ result = json.ResultString.FirstOrDefault(r => r.Step == step);
                        if (result != null)
                        {
                            sb.Append($"{result.Measured},");
                        }
                    }

                    swOut.WriteLine(sb.ToString());
                }
            }
        } catch (Exception ex)
        {
            // Handle any exceptions that may occur during file writing
            // Log or display the error message
        }
    }
    private void PrintJson2(string path, List<DataPackage> data) {
        try
        {
            using (StreamWriter swOut = new StreamWriter(path, true))
            {
                foreach (DataPackage dataSup in data)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append($"{dataSup.Asm},");
                    sb.Append($"{dataSup.Custommer},");
                    sb.Append($"{dataSup.DateTime},");
                    sb.Append($"{dataSup.Result},");

                    if (dataSup.Data.StartsWith("{Date:"))
                    {
                        //string failure = GetFailureValue(dataSup.Data);
                        //string decodedString = Regex.Unescape(failure);
                        //if (decodedString != "<>3.33" & decodedString != "<>3.35")
                        //{
                        //    continue;
                        //}

                        string bat = GetMeasuredValue(dataSup.Data, "2.1");
                        string crystal32k = GetMeasuredValue(dataSup.Data, "3.33");
                        string crystal8M = GetMeasuredValue(dataSup.Data, "3.35");
                        string current = GetMeasuredValue(dataSup.Data, "4.4");

                        sb.Append($"{bat},");
                        sb.Append($"{crystal32k},");
                        sb.Append($"{crystal8M},");
                        sb.Append($"{current},");
                    }
                    else
                    {
                        string[] split = dataSup.Data.Split(',');
                        //if (split[11] != " ; 3.33" & split[11] != " ; 3.33" & split[11] != " ; 3.35" & split[11] != " ; 3.35" & split[11] != "")
                        //{
                        //    continue;
                        //}
                        sb.Append($"{split[14]},");
                        sb.Append($"{split[54]},");
                        sb.Append($"{split[56]},");
                        sb.Append($"{split[64]},");
                    }

                    swOut.WriteLine(sb.ToString());
                }
            }
        } catch (Exception ex)
        {
            // Handle any exceptions that may occur during file writing
            // Log or display the error message
        }
    }
    private string GetMeasuredValue(string input, string step) {
        string pattern = @"Step:" + Regex.Escape(step) + @",Description:[^,]+,Tolerance:[^,]+,Measured:([^,]+)";
        Match match = Regex.Match(input, pattern);

        if (match.Success)
        {
            string measuredValue = match.Groups[1].Value;
            return measuredValue;
        }

        return null;
    }
    private string GetFailureValue(string input) {
        string pattern = @",Failure:([^,]+)";
        Match match = Regex.Match(input, pattern);

        if (match.Success)
        {
            string measuredValue = match.Groups[1].Value;
            return measuredValue;
        }

        return null;
    }
    private JsonConvert_ GetJson(string data) {
        try
        {
            data = data.Replace("},]", "}]");
            JsonConvert_ dataJson = JsonConvert.DeserializeObject<JsonConvert_>(data);
            return dataJson;
        } catch (JsonException ex)
        {
            // Handle any JSON deserialization exceptions here
            // Log or display the error message
            return null; // Return null or an appropriate default value in case of an error
        }
    }
    private List<DataPackage> SortDateTime(List<DataPackage> data) {
        List<DataPackage> dataList = new List<DataPackage>();

        foreach (DataPackage dataSup in data)
        {
            try
            {
                if (dataSup.Data.StartsWith("{Date:"))
                {
                    JsonConvert_ json = GetJson(dataSup.Data);
                    dataSup.DateTime = $"{json.Date},{json.Time}";
                }
                else
                {
                    string[] format = dataSup.Data.Split(',');
                    dataSup.DateTime = $"{format[0]},{format[1]}";
                }

                dataList.Add(dataSup);
            } catch (Exception ex)
            {
                // Handle any exceptions that may occur during parsing or data processing
                // Log or display the error message
            }
        }

        // Sort by DateTime
        try
        {
            dataList.Sort((data1, data2) => {
                DateTime dt1 = DateTime.ParseExact(data1.DateTime, "MM/dd/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
                DateTime dt2 = DateTime.ParseExact(data2.DateTime, "MM/dd/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
                return dt1.CompareTo(dt2);
            });
        } catch (Exception ex)
        {
            // Handle any exceptions that may occur during sorting
            // Log or display the error message
        }

        return dataList;
    }
    private List<DataPackage> SortDateTime2(List<DataPackage> data) {
        List<DataPackage> dataList = new List<DataPackage>();

        foreach (DataPackage dataSup in data)
        {
            try
            {
                if (dataSup.Data.StartsWith("{Date:"))
                {
                    string pattern = @"{Date:(\d{2}/\d{2}/\d{4}),Time:(\d{2}:\d{2}:\d{2})";
                    Match match = Regex.Match(dataSup.Data, pattern);
                    if (match.Success)
                    {
                        string date = match.Groups[1].Value; // Extracted date
                        string time = match.Groups[2].Value; // Extracted time
                        dataSup.DateTime = $"{date},{time}";
                    }
                    else
                    {
                        MessageBox.Show("Format json date time Error");
                    }
                }
                else
                {
                    string[] format = dataSup.Data.Split(',');
                    dataSup.DateTime = $"{format[0]},{format[1]}";
                }

                dataList.Add(dataSup);
            } catch (Exception ex)
            {
                // Handle any exceptions that may occur during parsing or data processing
                // Log or display the error message
            }
        }

        // Sort by DateTime
        try
        {
            dataList.Sort((data1, data2) => {
                DateTime dt1 = DateTime.ParseExact(data1.DateTime, "dd/MM/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
                DateTime dt2 = DateTime.ParseExact(data2.DateTime, "dd/MM/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
                return dt1.CompareTo(dt2);
            });
        } catch (Exception ex)
        {
            // Handle any exceptions that may occur during sorting
            // Log or display the error message
        }

        return dataList;
    }
    private void SortDate() {
        List<string> dates = new List<string>
        {
                "06/03/2023",
                "04/03/2023",
                "07/03/2023",
                "06/02/2023"
            };

        dates.Sort((date1, date2) => DateTime.Parse(date1).CompareTo(DateTime.Parse(date2)));
    }
    private void SortTime() {
        List<string> times = new List<string>
        {
                "01:36:53",
                "00:36:54",
                "12:36:53",
                "00:01:53"
            };

        times.Sort((time1, time2) => TimeSpan.Parse(time1).CompareTo(TimeSpan.Parse(time2)));
    }
    private void SortDateTime_() {
        List<string> dateTimeList = new List<string>
        {
                "07/03/2023,01:36:53",
                "06/03/2023,00:36:54",
                "06/03/2023,00:36:53",
                "01/03/2023,00:01:53"
            };

        dateTimeList.Sort((dateTime1, dateTime2) =>
        {
            DateTime dt1 = DateTime.ParseExact(dateTime1, "MM/dd/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime dt2 = DateTime.ParseExact(dateTime2, "MM/dd/yyyy,HH:mm:ss", CultureInfo.InvariantCulture);
            return dt1.CompareTo(dt2);
        });
    }
    private List<DataPackage> SortName(List<DataPackage> dataIn) {
        // Sort the list based on the 'Asm' property in ascending order
        List<DataPackage> sortedData = dataIn.OrderBy(dp => dp.Asm).ToList();

        return sortedData;
    }
    private List<DataPackage> GetdataShort(ref List<DataPackage> dataInPut) {
        List<DataPackage> dataOut = new List<DataPackage>();
        if (dataInPut.Count > 0)
        {
            string asm = dataInPut[0].Asm; // Get the Asm value of the first element
            int index = 0;
            while (index < dataInPut.Count && dataInPut[index].Asm == asm)
            {
                dataOut.Add(dataInPut[index]);
                index++;
            }
            dataInPut.RemoveRange(0, index); // Remove the submitted data from dataInPut
        }
        return dataOut;
    }
}

public class JsonConvert_ {
    public string Date { get; set; }
    public string Time { get; set; }
    public string LoginID { get; set; }
    public string SWVersion { get; set; }
    public string FWVersion { get; set; }
    public string SpecVersion { get; set; }
    public string TestTime { get; set; }
    public string LoadInOut { get; set; }
    public string Mode { get; set; }
    public string FinalResult { get; set; }
    public string SN { get; set; }
    public object Failure { get; set; }
    public List<ResultString_> ResultString { get; set; }

    public JsonConvert_() {
        Date = string.Empty;
        Time = string.Empty;
        LoginID = string.Empty;
        SWVersion = string.Empty;
        FWVersion = string.Empty;
        SpecVersion = string.Empty;
        TestTime = string.Empty;
        LoadInOut = string.Empty;
        Mode = string.Empty;
        FinalResult = string.Empty;
        SN = string.Empty;
        Failure = string.Empty;
        ResultString = new List<ResultString_>();
    }
    public class ResultString_ {
        public string Step { get; set; }
        public string Description { get; set; }
        public string Tolerance { get; set; }
        public string Measured { get; set; }
        public string Result { get; set; }

        public ResultString_() {
            Step = string.Empty;
            Description = string.Empty;
            Tolerance = string.Empty;
            Measured = string.Empty;
            Result = string.Empty;
        }
    }
}
