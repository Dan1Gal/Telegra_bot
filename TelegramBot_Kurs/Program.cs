using Telegram.Bot;
using System;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TelegramBot_Kurs
{

    class Program
    {

        static async Task Main(string[] args)
        {

            var botClient = new TelegramBotClient("*токен бота*");
            using var cts = new CancellationTokenSource();

            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }
            };

            botClient.StartReceiving(
                HandleUpdatesAsync,
                HandleErrorAsync,
                receiverOptions,
                cancellationToken: cts.Token);

            var me = await botClient.GetMeAsync();
            Console.ReadLine();

            cts.Cancel();

            




           
        }

static async Task HandleUpdatesAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
            {
                var message = update.Message;
                
                if (update.Type == UpdateType.Message && update?.Message?.Text != null)
                {
                    await HandleMessage(botClient, update.Message);
                    return;
                }

                if (update.Type == UpdateType.CallbackQuery)
                {
                    return;
                }

                


            }
 static Task HandleErrorAsync(ITelegramBotClient client, Exception exception, CancellationToken cancellationToken)
            {

                var ErrorMessage = exception switch
                {
                    ApiRequestException apiRequestException
                        => $"Ошибка телеграм АПИ:\n{apiRequestException.ErrorCode}\n{apiRequestException.Message}",
                    _ => exception.ToString()
                };
                Console.WriteLine(ErrorMessage);
                return Task.CompletedTask;
            }
        private static async Task HandleMessage(ITelegramBotClient botClient, Message message)
        {


            if (message.Text == "Привет" || message.Text == "привет")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Здравствуйте!\nНачнём работу?\nВыберите команды:\nВыбрать тему рекламы:/choose | Ссылка на бота: /link");
                return;
            }

            if (message.Text == "Пока" || message.Text == "пока")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Всего хорошего, буду с нетерпением ждать вашего возвращения:)\nЕсли захотите вернуться, просто нажмите сюда ----> /start");
                return;
            }

            if (message.Text != null)
            {
                Console.WriteLine($"{message.Chat.Username ?? "анон"}     |    {message.Text}");
            }


            if (message.Text == "/start")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Выберите команды:\n Выбрать тему рекламы:/choose | Ссылка на бота: /link");

                //работа с базой данных в Excel
                try
                {
                    using (BD.ExcelHelper helper = new BD.ExcelHelper())
                    {
                        if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                        {
                            for (int i = 2; i >= 10; i++)
                            {
                                if (Convert.ToString(helper.Get(column: "A", row: i)) == Convert.ToString(message.Chat.Id))
                                {
                                    i++;
                                    helper.Save();
                                }
                                if ((Convert.ToString(helper.Get(column: "A", row: i)) != Convert.ToString(message.Chat.Id)))
                                {
                                    i++;
                                    helper.Set(column: "A", row: i, data: Convert.ToString(message.Chat.Id));
                                    helper.Save();
                                }
                            }
                            helper.Save();
                        }
                    }

                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }

                //

                return;
            }

            if (message.Text == "/link")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "https://t.me/KursovReklama_bot");
                return;
            }

            if (message.Text == "/choose")
            {
                ReplyKeyboardMarkup keyboard = new(new[]
                {
            new KeyboardButton[] {"Реклама сайта", "Реклама соц. сети"},
            new KeyboardButton[] {"Реклама компании"}
        })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Выберите действия:", replyMarkup: keyboard);
                return;
            }

            //создание рекламы сайта по шаблону

            if (message.Text == "Реклама сайта")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите ссылку на ваш сайт используя формат: \nСайт - ссылка на сайт");
                return;
            }
            if (message.Text.StartsWith("Сайт"))
            {
                string rek = "Реклама!\n Привет дружище, нашёл для тебя сайт с очень интересным функционалом, ты точно не видел ещё ничего подобного. Скорее переходи по ссылке и удивляйся!\n" + message.Text;

                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }

            //создание рекламы компании по шаблону

            if (message.Text == "Реклама компании")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите название вашей компании используя формат:\n \"название компании\", чем занимается компания.");
                return;
            }
            if (message.Text.StartsWith("\""))
            {
                string rek = " Реклама!\n Привет дружище, если тебе нужна помощь по узкому профилю работ и ты давно ищешь такую компанию, что справилась бы с этой работой.То это именно то что ты искал!\n" + message.Text;

                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }

            //создание командных кнопок

            if (message.Text == "Реклама соц. сети")
            {
                ReplyKeyboardMarkup keyboard = new(new[]
                {
            new KeyboardButton[] {"Инстаграм","Телеграм"},
            new KeyboardButton[] {"Вконтакте","ТикТок"}
        })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Выберите соц. сеть:", replyMarkup: keyboard);
                return;
            }

            //создание рекламы Инстаграма по шаблону

            if (message.Text == "Инстаграм")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите ссылку на ваш Инстаграм аккаунт.");
                return;
            }
            if (message.Text.StartsWith("https://instagram") | message.Text.StartsWith("https://www.instagram"))
            {
                string rek = "Реклама!\n В Инстаграм уже год и за первые 4 месяца сделал 100 000 подписчиков, потом аккаунт получил теневой бан, но через месяц получил прирост еще в 90 000. Продолжаю снимать и монтировать, пишу об этом и делюсь секретами в своём профиле " + message.Text;

                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }

            //создание рекламы Телеграма по шаблону

            if (message.Text == "Телеграм")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите ссылку на вашу Телеграм группу.");
                return;
            }

            if (message.Text.StartsWith("@") | message.Text.StartsWith("https://t.me/"))
            {
                string rek = "Реклама!\n Создали с ребятами группу в Телеграм, где можно найти много интересного контента. Вы точно не сможете пройти мимо. Мы собираемся провести крупный розыгрыш на 10 000 подписчиков. Подпишись и участвуй в розыгрыше! Группа -" + message.Text;


                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }
            

            //создание рекламы ТикТока по шаблону

            if (message.Text == "ТикТок")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите ссылку на ваш ТикТок аккаунт.");
                return;
            }
            if (message.Text.StartsWith("https://www.tiktok.com") | message.Text.StartsWith("https://tiktok.com"))
            {
                string rek = "Реклама!\n Создали с ребятами ТикТок аккаунт, где можно найти много интересного контента.Вы точно не сможете пройти мимо. Мы собираемся провести крупный розыгрыш на 10 000 подписчиков.Подпишись и участвуй в розыгрыше! Аккаунт - " + message.Text;

                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }

            //создание рекламы Вконтакте по шаблону

            if (message.Text == "Вконтакте")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Введите ссылку на вашу группу Вконтакте.");
                return;
            }
            if (message.Text.StartsWith("https://vk.com") | message.Text.StartsWith("https://www.vk.com"))
            {
                string rek = "Реклама!\n В моей группе Вконтакте можно найти много интересного контента. Вы точно не сможете пройти мимо. Каждого подписчика считаю своим другом. Подпишись и давай творить вместе! Группа -" + message.Text;
                using (BD.ExcelHelper helper = new BD.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "BD.xlsx")))
                    {
                        for (int i = 2; i <= 3; i++)
                        {
                            if (Convert.ToString(helper.Get(column: "A", row: i)) != null)
                            {
                                string s = Convert.ToString(helper.Get("A", i));
                                Console.WriteLine(s);
                                await botClient.SendTextMessageAsync(s, rek);
                            }
                        }
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat.Id, "Реклама успешно разослана, спасибо что выбрали нас! :)");
                return;
            }

           



        }
    }
}