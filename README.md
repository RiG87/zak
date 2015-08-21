# Zimbra Attachments Killer (zak.py) 
# Microsoft Exchange Attachments Killer (meak.ps1)

`[zimbra_user@zimbra_mail_store]$ ./zak -h` <br>

`Zimbra Attachments Killer` <br>
`[-h] [-a ACCOUNTS] [-t TIME_TO_LIVE]` <br>
`[-s START] [-l LIMIT] [-d LOG_DIR]` <br>
`[--debug DEBUG] [--child CHILD] [-v]` <br>

Removes attachments from zimbra email messages <br>
Usages: <br>


* 
`[zimbra_user@mail_store]$ ./zak.py -a account@my.dom -t 180` <br>
will remove attachments older then 180 days in account@my.dom <br>


* 
`[zimbra_user@mail_store]$ cat list` <br>
`account1@my.dom` <br>
`account2@my.dom` <br>
`[zimbra_user@mail_store]$ ./zak.py -a list -t 180` <br>
will remove attachments older then 180 days <br>
in account1@my.dom and in account2@my.dom <br>

* 
`[zimbra_user@mail_store]$ ./zak.py -a all -t 180` <br>
will remove attachments older then 180 days in all accounts <br>


`optional arguments:` <br>
`-h, --help       show this help message and exit` <br>
`-a ACCOUNTS      list zimbra-accounts for processing. default: empty` <br>
`-t TIME_TO_LIVE  attachments time to live in days. default: 360` <br>
`-s START         account list Start position for processing. default: -1` <br>
`-l LIMIT         limit of accounts portion for processing. default: -1` <br>
`-d LOG_DIR       directory for log files. default: logs` <br>
`--debug DEBUG    is debug mode` <br>
`--child CHILD    is child process. Private flag` <br>
`-v, --version    show program's version number and exit` <br>



Для видалення вкладень з MS Exchange 2013 реалізовано powershell­script
`meak.ps1`. Для передачі параметрів запуску у скрипт, адміністратор повинен
використовувати powershell­синтаксис.

Приклад 1. Команда для видалення всіх файлів у скрині username:
`PS> .\meak.ps1 ­AllDatabases ­Account “username@my.dom” ­TimeToLive 0`

Приклад 2. Типова команда для запуску:
`PS> .\meak.ps1 ­AllDatabases ­AllAccounts ­TimeToLive 90 ­LogDir “D:\logs” ­Limit 100`

Як видно, для запуску потрібно визначити ряд параметрів:

<br>
* ­AllDatabases (flag) ­ Визначення бази даних. всі доступні бази даних
<br>
* ­Databases (list) ­ Визначення бази даних. одне, або декілька імен баз даних
<br>
Визначення списку акаунтів (обов’язково). Для визначення списку цільових акаунтів,
передбачено три взаємо виключаючі параметри (в порядку пріоритетності): 
<br>
* ­AllAccounts (flag) ­ всі доступні акаунти
<br>
* ­AccountsFile (file path) ­ шлях до файлу із списком акаунтів
<br>
* ­Account (string) ­ одна цільова email­адреса
<br>
* ­TimeToLive (integer) ­ Визначення часу життя вкладень (не обов’язково). За замовчуванням: 360 днів
<br>
* ­Limit (integer) ­ Визначення розміру порції (не обов’язково). За замовчуванням: 1000
<br>
* ­LogDir (string) ­ Визначення директорії для лог­файлів (не обов’язково). За замовчуванням: “.\logs”
<br>
* ­DebugMode (flag) ­ Визначення режиму роботи (не обов’язково). За замовчуванням відсутній
<br>