using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Yandex.Checkout.V3;

namespace TestUKassa
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

            var client = new Yandex.Checkout.V3.Client(
                    shopId: "815006",
                    secretKey: "test_b70xo7RIfIoq1u3LVkfNhZuARcKycBB-zlBCfrKDDjs");
            string PaymentId = "286386a2-000f-5000-9000-177e592577f9";
            Payment p = client.GetPayment(PaymentId);
            //var status = p.Status;
            //Console.WriteLine(DateTime.Now.ToString() + " ---- " + status.ToString());
            //while (status == p.Status)
            //{
            //    p = client.GetPayment(PaymentId);
            //    System.Threading.Thread.Sleep(5000);
            //}
            //Console.WriteLine(DateTime.Now.ToString() + " ---- " + p.Status.ToString());
            //Console.WriteLine("Press any key to stop...");
            //Console.ReadKey();
            //var p2 = client.GetPayment("");
            var newPayment = new NewPayment
            {
                Amount = new Amount { Value = 397.99m, Currency = "RUB" },
                Description = "Заказ №74",
                Capture = true,
                Metadata = new Dictionary<string, string>
                {
                    {"Номер","73" },
                    {"Менеджер","Ерохин" }
                },
                Confirmation = new Confirmation
                {
                    Type = ConfirmationType.Redirect,
                    ReturnUrl = "https://stinmarket.ru/"
                },
                Receipt = new Receipt
                {
                    Phone = "909-098",
                    Email = "valya.fl-63@yandex.ru",
                    Items = new List<ReceiptItem>
                    {
                        new ReceiptItem 
                        {
                            Description = "Наименование товара 1",
                            Quantity = 1.00m,
                            Amount = new Amount { Value = 250.00m, Currency = "RUB"},
                            VatCode = VatCode.NoVat,
                            PaymentMode = PaymentMode.FullPayment,
                            PaymentSubject = PaymentSubject.Commodity,
                        },
                        new ReceiptItem
                        {
                            Description = "Наименование товара 2",
                            Quantity = 4.00m,
                            Amount = new Amount { Value = 12.00m, Currency = "RUB"},
                            VatCode = VatCode.NoVat,
                            PaymentMode = PaymentMode.FullPayment,
                            PaymentSubject = PaymentSubject.Commodity,
                        },
                        new ReceiptItem
                        {
                            Description = "Наименование товара 3",
                            Quantity = 3.00m,
                            Amount = new Amount { Value = 33.33m, Currency = "RUB"},
                            VatCode = VatCode.NoVat,
                            PaymentMode = PaymentMode.FullPayment,
                            PaymentSubject = PaymentSubject.Commodity,
                        },
                    },
                    TaxSystemCode = TaxSystem.Simplified
                }
            };
            Payment payment = null;
            try
            {
                payment = client.CreatePayment(newPayment);
            }
            catch (YandexCheckoutException ye)
            {
                string sd = ye.Error.Description;
                sd += ye.Error.Parameter;
            }
            catch (Exception e)
            {
                string ss = "";
            }
            string url = payment.Confirmation.ConfirmationUrl;
            string rr = "";
        }
    }
}
