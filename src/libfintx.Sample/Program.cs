using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using CommandLine;
using libfintx.Data;

namespace libfintx.Sample
{
    class Program
    {
        enum OperationType
        {
            Accounts,
            Balance,
            Transactions,
            AllBalances
        }

        class Options
        {
            [Option('b', "blz", Required = true, HelpText = "Set the FinTS bank code.")]
            public int BankCode { get; set; }

            [Option('d', "bic", Required = true, HelpText = "Set the FinTS BIC.")]
            public string Bic { get; set; }

            [Option('i', "userid", Required = true, HelpText = "Set the FinTS user id.")]
            public string UserId { get; set; }

            [Option('a', "accountnumber", Required = false, Default = "", HelpText = "Set the FinTS bank account number. Not required for operation==AllBalances")]
            public string Account { get; set; }

            [Option('u', "url", Required = true, HelpText = "Set the FinTS URL.")]
            public string Url { get; set; }

            [Option('p', "pin", Required = true, HelpText = "Set the FinTS PIN.")]
            public string Pin { get; set; }

            [Option('o', "operation", Required = true, HelpText = "Set the FinTS transaction type.")]
            public OperationType Operation { get; set; }
        }

        static async Task<string> WaitForTanAsync(TANDialog tanDialog)
        {
            foreach (var msg in tanDialog.DialogResult.Messages)
                Console.WriteLine(msg);

            return await Task.FromResult(Console.ReadLine());
        }

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            Parser.Default.ParseArguments<Options>(args)
            .WithParsed(o =>
            {
                var details = new ConnectionDetails()
                {
                    Url = o.Url,
                    Account = o.Account,
                    Blz = o.BankCode,
                    Bic = o.Bic,
                    Pin = o.Pin,
                    UserId = o.UserId,
                };

                var client = new FinTsClient(details);
                switch (o.Operation)
                {
                    case OperationType.Accounts:
                        Accounts(client);
                        break;
                    case OperationType.Balance:
                        Balance(client);
                        break;
                    case OperationType.Transactions:
                        Transactions(client);
                        break;
                    case OperationType.AllBalances:
                        AllBalances(client);
                        break;
                }
            });

            Console.ReadLine();
        }

        private static async void Accounts(FinTsClient client)
        {
            var result = await client.Accounts(new TANDialog(WaitForTanAsync));
            if (!result.IsSuccess)
            {
                HBCIOutput(result.Messages);
                return;
            }

            Console.WriteLine("Account count: {0}", result.Data.Count);
            foreach (var account in result.Data)
            {
                Console.WriteLine("Account - Holder: {0}, Number: {1}", account.AccountOwner, account.AccountNumber);
            }
        }

        private static async void Balance(FinTsClient client)
        {
            var result = await client.Balance(new TANDialog(WaitForTanAsync));
            if (!result.IsSuccess)
            {
                HBCIOutput(result.Messages);
                return;
            }

            Console.WriteLine("Balance is: {0}\u20AC", result.Data.Balance);
        }

        private static async void Transactions(FinTsClient client)
        {
            var result = await client.Transactions(new TANDialog(WaitForTanAsync));
            if (!result.IsSuccess)
            {
                HBCIOutput(result.Messages);
                return;
            }

            Console.WriteLine("Transaction count:", result.Data.Count);
            foreach (var trans in result.Data)
            {
                Console.WriteLine("Transaction - Start Date: {0}, Amount: {1}\u20AC", trans.StartDate, trans.EndBalance - trans.StartBalance);
            }
        }

        private static async void AllBalances(FinTsClient client)
        {
            var result = await client.Accounts(new TANDialog(WaitForTanAsync));
            if (!result.IsSuccess)
            {
                HBCIOutput(result.Messages);
                return;
            }

            foreach (var account in result.Data)
            {
                client.activeAccount = account;
                if (account.AccountPermissions.Exists(x => x.Segment == "HKSAL"))
                {
                    var balance = await client.Balance(new libfintx.TANDialog(WaitForTanAsync));
                    if (!balance.IsSuccess)
                    {
                        HBCIOutput(balance.Messages);
                        Console.WriteLine("Account - Balance:           | Holder: {0,-32} | Number: {1,12} | Type: {2} | Error when trying to get Balance", account.AccountOwner, account.AccountNumber);
                    }
                    else
                    {
                        Console.WriteLine("Account - Balance: {0,8}\u20AC | Holder: {1,-32} | Number: {2,12} | Type: {3}", balance.Data.Balance, account.AccountOwner, account.AccountNumber, account.AccountType);
                    }
                } else
                {
                    Console.WriteLine("Account -                    | Holder: {0,32} | Number: {1,12} | Type: {2} | Balance not possible for this Account", account.AccountOwner, account.AccountNumber);
                }
            }
        }

        /// <summary>
        /// HBCI-Nachricht ausgeben
        /// </summary>
        /// <param name="hbcimsg"></param>
        private static void HBCIOutput(IEnumerable<HBCIBankMessage> hbcimsg)
        {
            foreach (var msg in hbcimsg)
            {
                Console.WriteLine("Code: " + msg.Code + " | " + "Typ: " + msg.Type + " | " + "Nachricht: " + msg.Message);
            }
        }
    }
}
