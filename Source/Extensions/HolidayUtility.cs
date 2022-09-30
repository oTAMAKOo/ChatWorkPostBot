using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Extensions
{
    public static class HolidayUtility
    {
        private const string HolidayCsvUrl = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv";

        public static async Task<Dictionary<DateTime, string>> GetHoliday()
        {
            var holidayDictionary = new Dictionary<DateTime, string>();

            // .Net5でSJISを使う場合に必要.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // CSVをダウンロード.
            var client = new System.Net.WebClient();
            
            var buffer = await client.DownloadDataTaskAsync(HolidayCsvUrl);

            var str = Encoding.GetEncoding("shift_jis").GetString(buffer);
 
            // 行毎に配列に分割.
            var rows = str.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
 
            // 一行目を飛ばしてデータをディクショナリに格納.
            for (var i = 1; i < rows.Length; i++)
            {
                var cols = rows[i].Split(',');

                holidayDictionary.Add(DateTime.Parse(cols[0]), cols[1]);
            }

            return holidayDictionary;
        }
    }
}
