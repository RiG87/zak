# Zimbra Attachments Killer

`[zimbra_user@zimbra_mail_store]$ ./zak -h` <br>

`Zimbra Attachments Killer` <br>
`[-h] [-a ACCOUNTS] [-t TIME_TO_LIVE]` <br>
`[-s START] [-l LIMIT] [-d LOG_DIR]` <br>
`[--debug DEBUG] [--child CHILD] [-v]` <br>

Removes attachments from zimbra email messages <br>
Usages: <br>

1. <br>
`[zimbra_user@mail_store]$ ./zak.py -a account@my.dom -t 180` <br>

will remove attachments older then 180 days in account@my.dom <br>

2. <br>
`[zimbra_user@mail_store]$ cat list` <br>
account1@my.dom <br>
account2@my.dom <br>
`[zimbra_user@mail_store]$ ./zak.py -a list -t 180` <br>

will remove attachments older then 180 days <br>
in account1@my.dom and in account2@my.dom <br>

3. <br>
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