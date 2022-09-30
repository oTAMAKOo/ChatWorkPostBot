
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Extensions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace ChatWorkPostBot
{
    public sealed class SpreadsheetService : Singleton<SpreadsheetService>
    {
        //----- params -----

        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        //----- field -----

        private SheetsService sheetsService = null;

        //----- property -----
        
        //----- method -----

        private SpreadsheetService(){ }

        public Task Initialize()
        {
            sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = ConnectOAuth(),
                ApplicationName = "ChatWorkPostBot",
            });

            return Task.CompletedTask;
        }

        /// <summary> Spread Sheetにアクセストークンを取得して接続 </summary>
        private UserCredential ConnectOAuth()
        {
            var setting = Setting.Instance;

            var oauthKeyPath = setting.GetOAuthKeyPath();

            UserCredential credential;

            using (var stream = new FileStream(oauthKeyPath, FileMode.Open, FileAccess.Read))
            {
                var credPath = "OAuthToken";

                var secrets = GoogleClientSecrets.Load(stream).Secrets;

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;

                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        public async Task<ValueRange> GetValue(string range)
        {
            var setting = Setting.Instance;

            var spreadsheetId = setting.SpreadsheetId;

            var req = sheetsService.Spreadsheets.Values.Get(spreadsheetId, $"{setting.SheetName}!{range}");

            var result = await req.ExecuteAsync();

            return result;
        }
    }
}
