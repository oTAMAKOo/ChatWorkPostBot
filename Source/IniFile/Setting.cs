
using System;
using Extensions;
using IniFileParser.Model;

namespace ChatWorkPostBot
{
    public sealed class Setting : IniFile<Setting>
    {
        //----- params -----

        #region Chatwork

        private const string ChatworkSection = "Chatwork";
        
        private const string ChatworkApiKeyField = "ApiKey";

        #endregion

        #region Spreadsheet

        private const string SpreadsheetSection = "Spreadsheet";

        private const string SpreadsheetIdField = "SpreadsheetId";

        private const string SpreadsheetNameField = "SheetName";

        private const string OAuthKeyField = "OAuthKey";

        #endregion

        #region PostData

        private const string PostDataSection = "PostData";

        private const string DataStartRowField = "DataStartRow";

        private const string MessageColumnField = "MessageColumn";

        private const string RoomIdColumnField = "RoomIdColumn";
        
        private const string PostIntervalColumnField = "PostIntervalColumn";
        
        private const string PostDayColumnField = "PostDayColumn";

        private const string PostTimeColumnField = "PostTimeColumn";

        private const string PostHolidayColumnField = "PostHoliday";

        #endregion

        //----- field -----
        
        //----- property -----

        public override string FileName { get { return "setting"; } }
        
        public string ChatworkApiKey { get { return GetData<string>(ChatworkSection, ChatworkApiKeyField); } }

        public string SpreadsheetId { get { return GetData<string>(SpreadsheetSection, SpreadsheetIdField); } }

        public string SheetName { get { return GetData<string>(SpreadsheetSection, SpreadsheetNameField); } }

        public int DataStartRow { get { return GetData<int>(PostDataSection, DataStartRowField); } }

        public string MessageColumn { get { return GetData<string>(PostDataSection, MessageColumnField); } }

        public string RoomIdColumn { get { return GetData<string>(PostDataSection, RoomIdColumnField); } }

        public string PostIntervalColumn { get { return GetData<string>(PostDataSection, PostIntervalColumnField); } }

        public string PostDayColumn { get { return GetData<string>(PostDataSection, PostDayColumnField); } }

        public string PostTimeColumn { get { return GetData<string>(PostDataSection, PostTimeColumnField); } }

        public string PostHolidayColumn { get { return GetData<string>(PostDataSection, PostHolidayColumnField); } }

        //----- method -----

        protected override void OnLoad()
        {
            Console.WriteLine($"Setting : Load { GetConfigFilePath() }");
        }

        protected override void SetDefaultData(ref IniData data)
        {
            // Chatwork.
            data[ChatworkSection][ChatworkApiKey] = "ABCDEFG123456789";
            
            // Spreadsheet.
            data[SpreadsheetSection][SpreadsheetIdField] = "1234567890";
            data[SpreadsheetSection][SpreadsheetNameField] = "1234567890";
            data[SpreadsheetSection][OAuthKeyField] = "client_secret.json";

            // PostData.
            data[PostDataSection][DataStartRowField] = "1";
            data[PostDataSection][MessageColumnField] = "A";
            data[PostDataSection][RoomIdColumnField] = "B";
            data[PostDataSection][PostIntervalColumnField] = "C";
            data[PostDataSection][PostDayColumnField] = "D";
            data[PostDataSection][PostTimeColumnField] = "E";
            data[PostDataSection][PostHolidayColumnField] = "F";
        }

        public string GetOAuthKeyPath()
        {
            var configFileDirectory = ConfigUtility.GetConfigFolderDirectory();

            var fileName = GetData<string>(SpreadsheetSection, OAuthKeyField);

            var oauthKeyPath = PathUtility.Combine(configFileDirectory, fileName);

            return oauthKeyPath;
        }
    }
}
