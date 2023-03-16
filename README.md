# Telegram PowerPoint Server
Программа, которая с помощью Telegram Bot API взаимодействует с Microsoft PowerPoint. Позволяет удобно проводить демонстрацию презентации с помощью кнопок в чате Telegram.

## О проекте.
> Так как есть люди, которые занимаются **научной** деятельностью и **часто выступают**, им было бы наверняка удобнее переключать слайды через **бота в Telegram** через **кнопки под сообщением**.
> Эта программа **позволяет реализовать** это в жизни.
> Единственное, что нужно сделать - **сгенерировать токен** в [BotFather](t.me/BotFather) и **иметь хорошее подключение к Интернету**.

## О компиляции.
> Разрабатывалось и тестировалось на платформе AutoHotkey (**версия 1.1.35.00**).
> Рекомендуется производить компиляцию именно на этой версии.
> После компиляции программа обрабатывалась через VMProtect с целью сокрытия исходного кода.

## О том, где тестировалось.
> Тестировалось на **Windows 10** и **Windows 11** в **Microsoft PowerPoint 2016**.

## О том, как получить токен.
> 1. Откройте чат с [BotFather](t.me/BotFather).
> 2. Отправьте боту сообщение: **/newbot**.
> 3. Отправьте боту сообщение с **названием** Вашего нового бота.
> 4. Отправьте боту сообщение с **username'ом** Вашего нового бота (он будет использоваться в ссылке на Вашего бота).
> 5. Во втором абзаце после текста «**Use this token to access the HTTP API:**» скопируйте токен. Обычно он в формате **ХХХХХХХХ:ХХХХХХХХХХХХХХ**...
> 6. **Вставьте в программу**.
>
> Введенный Вами токен **никуда не передается** и сохраняется только на **Вашем** устройстве.

## О том, как пользоваться.
> Чтобы запустить работу сервера, **сгенерируйте токен** по инструкции указанной выше и **вставьте в поле для ввода** токен, и нажмите кнопку «**Продолжить**».
> Если Вы указали все верно, то **программа отправит в трее уведомление с кодом**, который **нужно будет отправить боту**.
> После отправки кода, бот должен Вам **ответить** тем, что **он успешно привязан** к Вашему аккаунту.
> С этого момента **управление будет проходить в чате**, оно **интуитивно понятное**. Просто **придерживайтесь инструкций** от бота.
> Если Вам нужно будет **отключить программу**, то найдите **значок программы в трее** на панеле задач и **кликните по нему правой кнопкой мыши** и выберите пункт «**Отключить сервер**».

## О том, как компилировать Telegrma PowerPoint Server.
> 1. Скачайте файл «ppts.ahk» на Ваше устройство.
> 2. Установите AutoHotkey последней версии (**1.\***), но рекомендуется использовать версию «**1.1.35.00**».
> 3. Откройте «**ahk2exe.exe**» через меню «**Выполнить**» (**Win + R**).
> 4. В поле «**Source (script file)**» укажите путь к скачанному Вами файлу.
> 5. В поле «**Base file (.bin, .exe)**» выберите вариант содержащий слово «**ANSI**» (**в разных версиях по-разному**).
> 6. Затем нажмите «**Convert**» или «**Compile**» (**в разных версиях по-разному**).
> 7. **Готово**. Ваш скомпилированный файл расположен в той же директории, где и хранится скачанный Вами файл (**если Вы не изменяли поле «Destination (.exe file)»**).
