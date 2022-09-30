
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Extensions;
using Newtonsoft.Json;

namespace ChatWorkPostBot
{
    public sealed class PostTriggerService : Singleton<PostTriggerService>
    {
        //----- params -----

        private const int PostDataFetchInterval = 5;

        //----- field -----

        // 休日一覧.
        private Dictionary<DateTime, string> holidays = null;
        // 最終投稿時間.
        private Dictionary<int, DateTime?> postHistory = null;
        // 投稿データ.
        private PostData[] postData = null;
        // 最終投稿データ取得時間.
        private DateTime? lastFetchTime = null;

        //----- property -----
        
        //----- method -----

        private PostTriggerService(){ }

        public async Task Initialize()
        {
            postHistory = new Dictionary<int, DateTime?>();

            // 休日情報取得.
            
            try
            {
                holidays = await HolidayUtility.GetHoliday();

                Console.WriteLine("PostTriggerService : Get holiday data.\n");
            }
            catch (Exception e)
            {
                Console.WriteLine($"PostTriggerService : {e}\n");

                throw;
            }
        }

        public async Task Update(CancellationToken cancelToken)
        {
            var now = DateTime.Now;

            // 投稿情報取得.

            if (!lastFetchTime.HasValue || lastFetchTime.Value.AddMinutes(PostDataFetchInterval) < now)
            {
                postData = await GetPostData();
                
                lastFetchTime = now;
            }

            if (postData == null){ return; }

            // 投稿.

            var tasks = new List<Task>();

            foreach (var data in postData)
            {
                var dataHash = data.GetHashCode();

                if (string.IsNullOrEmpty(data.roomId)){ continue; }

                try
                {
                    if (!CheckPostTime(now, data, dataHash)){ continue; }

                    var task = PostMessage(now, data, dataHash, cancelToken);

                    tasks.Add(task);
                }
                catch (Exception e)
                {
                    await SendMessage(data.roomId, e.ToString(), cancelToken);

                    postHistory[dataHash] = now;
                }
            }

            await Task.WhenAll(tasks);
        }

        private async Task PostMessage(DateTime now, PostData data, int dataHash, CancellationToken cancelToken)
        {
            var setting = Setting.Instance;

            var client = new ChatworkClient(data.roomId, setting.ChatworkApiKey);

            await client.SendMessage(data.message, cancelToken);

            postHistory[dataHash] = now;

            var json = JsonConvert.SerializeObject(data, Formatting.Indented);

            Console.WriteLine($"Post:\n{json}\n");
        }

        private async Task SendMessage(string roomId, string message, CancellationToken cancelToken)
        {
            var setting = Setting.Instance;

            var client = new ChatworkClient(roomId, setting.ChatworkApiKey);

            await client.SendMessage(message, cancelToken);
        }

        private bool CheckPostTime(DateTime now, PostData data, int dataHash)
        {
            // 祝日判定.

            if (!data.postHoliday)
            {
                if (holidays.Keys.Any(x => x.Date == now.Date)){ return false; }
            }

            // 投稿履歴判定.
            
            var lastPostTime = postHistory.GetValueOrDefault(dataHash);

            if (lastPostTime.HasValue)
            {
                // 同じ日には投稿しない.

                if (lastPostTime.Value.Month == now.Month && lastPostTime.Value.Day == now.Day) { return false; }

                // 投稿間隔判定.
                
                var intervalTime = lastPostTime.Value.Date.AddDays(data.dayInterval) + data.postTime;

                if (now < intervalTime){ return false; }
            }

            // 曜日判定.

            if (data.dayOfWeeks.All(x => x != now.DayOfWeek)){ return false; }

            // 時間判定.

            var postTime = data.postTime;

            if(now.Hour != postTime.Hours || now.Minute != postTime.Minutes || now.Second < postTime.Seconds) { return false; }

            return true;
        }

        private async Task<PostData[]> GetPostData()
        {
            var setting = Setting.Instance;

            var spreadsheetService = SpreadsheetService.Instance;

            var valueRange = await spreadsheetService.GetValue("A:Z");

            var list = new List<PostData>();

            for (var i = 0; i < valueRange.Values.Count; i++)
            {
                if (i < setting.DataStartRow){ continue; }

                try
                {
                    var item = valueRange.Values[i];

                    var roomIdValue = GetValue<string>(item, setting.RoomIdColumn);
                    var messageValue = GetValue<string>(item, setting.MessageColumn);
                    var dayIntervalValue = GetValue<int>(item, setting.PostIntervalColumn);
                    var dayOfWeeksValue = GetValue<string>(item, setting.PostDayColumn);
                    var postTimeValue = GetValue<string>(item, setting.PostTimeColumn);
                    var postHolidayValue = GetValue<bool>(item, setting.PostHolidayColumn);
                
                    var postData = new PostData()
                    {
                        roomId = roomIdValue,
                        message = messageValue,
                        dayInterval = dayIntervalValue,
                        dayOfWeeks = ConvertToDayOfWeeks(dayOfWeeksValue),
                        postTime = TimeSpan.Parse(postTimeValue),
                        postHoliday = postHolidayValue,
                    };

                    list.Add(postData);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }

            return list.ToArray();
        }

        private T GetValue<T>(IList<object> list, string column)
        {
            var index = ExcelUtility.ToAlphabetIndex(column) - 1;

            var value = list.ElementAtOrDefault(index);

            if (value == null){ return default; }

            return (T)Convert.ChangeType(value, typeof(T)); 
        }

        private DayOfWeek[] ConvertToDayOfWeeks(string dayOfWeeksValue)
        {
            var list = new List<DayOfWeek>();

            var parts = dayOfWeeksValue.Split(',').Select(x => x.Trim()).ToArray();

            var allDay = new DayOfWeek[]
            {
                DayOfWeek.Monday, 
                DayOfWeek.Tuesday, 
                DayOfWeek.Wednesday, 
                DayOfWeek.Thursday, 
                DayOfWeek.Friday, 
                DayOfWeek.Saturday, 
                DayOfWeek.Sunday
            };

            foreach (var item in parts)
            {
                switch (item)
                {
                    // 日本語指定.
                    case "月":  list.Add(DayOfWeek.Monday); break;
                    case "火":  list.Add(DayOfWeek.Tuesday); break;
                    case "水":  list.Add(DayOfWeek.Wednesday); break;
                    case "木":  list.Add(DayOfWeek.Thursday); break;
                    case "金":  list.Add(DayOfWeek.Friday); break;
                    case "土":  list.Add(DayOfWeek.Saturday); break;
                    case "日":  list.Add(DayOfWeek.Sunday); break;
                    case "全日": allDay.ForEach(x => list.Add(x)); break;

                    // 英語(短縮)指定.
                    case "Mon": list.Add(DayOfWeek.Monday); break;
                    case "Tue": list.Add(DayOfWeek.Tuesday); break;
                    case "Wed": list.Add(DayOfWeek.Wednesday); break;
                    case "Thu": list.Add(DayOfWeek.Thursday); break;
                    case "Fri": list.Add(DayOfWeek.Friday); break;
                    case "Sat": list.Add(DayOfWeek.Saturday); break;
                    case "Sun": list.Add(DayOfWeek.Sunday); break;

                    case "Mo": list.Add(DayOfWeek.Monday); break;
                    case "Tu": list.Add(DayOfWeek.Tuesday); break;
                    case "We": list.Add(DayOfWeek.Wednesday); break;
                    case "Th": list.Add(DayOfWeek.Thursday); break;
                    case "Fr": list.Add(DayOfWeek.Friday); break;
                    case "Sa": list.Add(DayOfWeek.Saturday); break;
                    case "Su": list.Add(DayOfWeek.Sunday); break;

                    case "All": allDay.ForEach(x => list.Add(x)); break;
                }
            }

            return list.ToArray();
        }
    }
}
