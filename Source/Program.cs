using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Extensions;

namespace ChatWorkPostBot
{
    public class Program
    {
        static async Task<int> Main(string[] args)
        {
            var exitCode = 0;

            Console.OutputEncoding = Encoding.UTF8;

            try
            {
                var cancelSource = new CancellationTokenSource();

                CheckExit(cancelSource).Forget();

                var mainHub = new MainHub();

                await mainHub.Initialize();

                while (!cancelSource.IsCancellationRequested)
                {
                    await mainHub.Update(cancelSource.Token);

                    await Task.Delay(TimeSpan.FromSeconds(5), cancelSource.Token);
                }
            }
            catch (TaskCanceledException)
            {
                /* Canceled for exit */
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync(ex.ToString());

                exitCode = Marshal.GetHRForException(ex);

                Console.ReadKey();
            }

            return exitCode;
        }

        private static async Task CheckExit(CancellationTokenSource cancelSource)
        {
            while (!cancelSource.IsCancellationRequested)
            {
                // Escキーで終了.

                if(Console.KeyAvailable)
                {
                    if(Console.ReadKey().Key == ConsoleKey.Escape)
                    {
                        cancelSource.Cancel();

                        Environment.Exit(0);
                    }
                }

                await Task.Delay(1, cancelSource.Token);
            }
        }
    }
}
