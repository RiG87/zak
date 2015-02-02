# Zimbra Attachments Killer

`[zimbra_user@zimbra_mail_store]$ ./zak -h`

usage: Zimbra Attachments Killer [-h] [-a ACCOUNTS] [-t TIME_TO_LIVE]
                                 [-s START] [-l LIMIT] [-d LOG_DIR]
                                 [--debug DEBUG] [--child CHILD] [-v]

Removes attachments from zimbra email messages

Usages:

### 1

`[zimbra_user@mail_store]$ ./zak.py -a account@my.dom -t 180`

    will remove attachments older then 180 days in account@my.dom

### 2

`[zimbra_user@mail_store]$ cat list`
account1@my.dom
account2@my.dom

`[zimbra_user@mail_store]$ ./zak.py -a list -t 180`

    will remove attachments older then 180 days
    in account1@my.dom and in account2@my.dom

### 3

`[zimbra_user@mail_store]$ ./zak.py -a all -t 180`

    will remove attachments older then 180 days in all accounts

optional arguments:
  -h, --help       show this help message and exit
  -a ACCOUNTS      list zimbra-accounts for processing. default: empty
  -t TIME_TO_LIVE  attachments time to live in days. default: 360
  -s START         account list Start position for processing. default: -1
  -l LIMIT         limit of accounts portion for processing. default: -1
  -d LOG_DIR       directory for log files. default: logs
  --debug DEBUG    is debug mode.
  --child CHILD    is child process. Private flag.
  -v, --version    show program's version number and exit