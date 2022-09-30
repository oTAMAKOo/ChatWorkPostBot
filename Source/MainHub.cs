
using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace ChatWorkPostBot
{
    public class MainHub
    {
        //----- params -----
        
        //----- field -----

        //----- property -----

        //----- method -----

        public async Task Initialize()
        {
            Console.WriteLine("\n------ Initialize ----------------\n");

            // IniFiles.
            
            var setting = Setting.CreateInstance();

            await setting.Load();

            // SSL.

            ServicePointManager.ServerCertificateValidationCallback = OnRemoteCertificateValidationCallback;

            // Spreadsheet.

            var spreadsheetService = SpreadsheetService.CreateInstance();

            await spreadsheetService.Initialize();

            // PostTrigger.

            var postTriggerService = PostTriggerService.CreateInstance();

            await postTriggerService.Initialize();

            ConsoleUtility.Separator();
        }

        public async Task Update(CancellationToken cancelToken)
        {
            var postTriggerService = PostTriggerService.Instance;

            try
            {
                await postTriggerService.Update(cancelToken);
            }
            catch (Exception e)
            {
                ConsoleUtility.Separator();

                Console.WriteLine(e);
                
                ConsoleUtility.Separator();
            }
        }

        // 信頼できないSSL証明書を「問題なし」にするメソッド
        private bool OnRemoteCertificateValidationCallback(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true; // 「SSL証明書の使用は問題なし」と示す
        }
    }
}
