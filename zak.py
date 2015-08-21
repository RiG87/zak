#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Zimbra Attachments Killer
# Automation script. Removes attachments from zimbra.
#

__program__ = 'Zimbra Attachments Killer'
__version__ = '1.2'
__author__ = 'pysarenko.a'

import os
import sys
import logging
import MySQLdb
import subprocess

from datetime import datetime
from datetime import timedelta

from time import strftime
from time import mktime

from email import message_from_file
from email.utils import parsedate
from email.mime.text import MIMEText
from email.generator import Generator

from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

from errno import EEXIST
from getpass import getuser
from contextlib import closing

from textwrap import dedent
from StringIO import StringIO

from ConfigParser import ConfigParser

# <editor-fold desc="Constants">
###############################################################################

ZIMBRA_CONFIG_CMD = ['/opt/zimbra/bin/zmlocalconfig', '-s']
ZIMBRA_CONFIG_CMD_doc = """shell cmd for retrieve zimbra config"""

ZIMBRA_STORE_DIR = '/srv/mail/store/'
ZIMBRA_STORE_DIR_doc = """path to zimbra mail store"""

_tmp_block = 'main'             # tmp config block need for ConfigParser()

_def_start = -1                 # default sql start
_def_limit = -1                 # default sql limit
_def_ttl_days = 360             # default files time to live in days
_def_all_accounts_flag = '*'    # default flag 'use full accounts list for processing'
_def_accounts = ''              # default accounts val
_def_log_dir = 'logs'           # default log-file name
_def_child = str(-1)            # default child counter

_zcp_home = 'zimbra_home'               # zimbra config param 'zimbra_home'
_zcp_user = 'zimbra_user'               # zimbra config param 'zimbra_user'
_zcp_db_dir = 'zimbra_db_directory'     # zimbra config param 'zimbra_db_directory'
_zcp_db_host = 'mysql_bind_address'     # zimbra config param 'mysql_bind_address'
_zcp_db_user = 'zimbra_mysql_user'      # zimbra config param 'zimbra_mysql_user'
_zcp_db_pass = 'zimbra_mysql_password'  # zimbra config param 'zimbra_mysql_password'
_zcp_db_socket = 'mysql_socket'         # zimbra config param 'mysql_socket'

###############################################################################
# </editor-fold>

# <editor-fold desc="Code Base">
###############################################################################


class Utils:
    """
    Some utils
    """

    log = None
    log_format = u'[%(levelname)-5s] [%(asctime)s] [line:%(lineno)-3s] %(message)s'

    @staticmethod
    def ensure_dir(full_name):
        """
        Ensure that a named directory exists; if it does not, attempt to create it.
        """
        try:
            os.makedirs(full_name)
        except OSError, e:
            if e.errno != EEXIST:
                raise

    @staticmethod
    def array_to_string(src_list=None, delimiter=', ', quoted=True, delimited=True):
        """
        Makes quoted string src_list

        :param src_list: list of values for build string
        :param delimiter: result string delimiter
        :param quoted: if True - value will be quoted
        :param delimited: if True - delimiter will be used
        :return: str
        """
        if len(src_list) > 0:
            res = ''
            for val in src_list:
                if quoted:
                    val = "'"+str(val)+"'"
                if delimited:
                    val = "%s%s" % (delimiter, str(val))
                res += val
                pass

            if delimited:
                res = res.replace(delimiter, '', 1)

            return res
        return ''

    @staticmethod
    def ex_call(cmd, logger=None):
        """
        Calls external shell command with params 'cmd', redirects stderr to the logfile.
        :returns: stdout
        """

        if logger:
            logger.debug("cmd " + Utils.array_to_string(src_list=cmd))

        if not cmd:
            Utils.error_exit("got empty cmd, when call %s" % cmd, logger=logger)

        p = subprocess.Popen(cmd, stderr=subprocess.STDOUT, stdout=subprocess.PIPE)
        (out, err) = p.communicate()

        if err:
            Utils.error_exit("got %s, when call %s" % (err, cmd), logger=logger)

        sys.stdout.flush()
        return out

    @staticmethod
    def get_logger(log_file, debug=False):
        logger = logging.getLogger(__program__)
        logger.setLevel(logging.getLevelName('DEBUG' if debug else 'INFO'))
        _logFileHandler = logging.FileHandler(log_file)
        _logFileHandler.setFormatter(logging.Formatter(Utils.log_format))
        logger.addHandler(_logFileHandler)
        return logger

    @staticmethod
    def error_exit(msg, logger=None):
        """
        Exit from script with error.
        """
        if logger:
            logger.error(msg)
        print
        print msg
        sys.exit(1)
        pass

    def __init__(self):
        """
        It's static class. Nothing to do.
        """
        pass
    pass


class Account:
    file_query = "SELECT concat('%s',(mailbox_id >> 12),'/',mailbox_id,'/msg/',(id >> 12),'/',id,'-',mod_content,'.msg') AS file " \
                 "FROM mboxgroup%s.mail_item where mailbox_id='%s'"

    file_query_doc = """
    MySQL query, which construct messages file names from zimbra metadata

    http://wiki.zimbra.com/wiki/Account_mailbox_database_structure

    params:
        %s store_dir - path to zimbra storage,
        %s group_id - id of mailboxgroup,
        %s mailbox_id - id of users mailbox
    """

    def delete_attachments(self, ttl, log, debug=True):
        """
        Removes attachments from message files for this account

        :param ttl: attachments time to live in days
        :param log: logger
        :param debug: if true - files will not be overwritten
        :return: log messages
        """
        for path in self.message_paths:
            if not os.path.isfile(path):
                log.debug("continue: message file not exists")
                continue

            # https://docs.python.org/3/library/email-examples.html
            with open(path) as fp:
                msg = message_from_file(fp)

                _date = parsedate(msg['Date'])
                if not _date:
                    log.debug('continue: message date header not exists')
                    continue

                _str_date = strftime("%d.%m.%Y %H:%M:%S", _date)
                _trg_date = datetime.now() - timedelta(days=(ttl + 1))

                if datetime.fromtimestamp(mktime(_date)) >= _trg_date:
                    log.debug('continue: too new ' + _str_date)
                    continue

                log.info('Message: ' + path)
                need_rewrite = False
                for part in msg.walk():
                    _content_maintype = part.get_content_maintype()
                    _content_type = part.get_content_type()

                    if _content_maintype == 'multipart' \
                            or _content_maintype == 'message' \
                            or _content_type == 'text/html' \
                            or _content_type == 'text/plain':
                        log.debug('Not Deleted. %s' % str(_content_type))
                        continue
                    else:
                        """
                        Delete attachment from message
                        Last chance for save file to disk.
                        """

                        log.info("Deleted. MsgDate: %s, Size: %s" % (_str_date, len(part.get_payload())))
                        part.set_payload('')
                        need_rewrite = True
                    pass
                pass
            pass

            if not debug and need_rewrite:
                """
                Write new message without attachments to disk
                """
                out_file = open(path, 'w')
                generator = Generator(out_file)
                generator.flatten(msg)
                out_file.close()
            else:
                if debug:
                    log.debug("continue: debug mode")
                else:
                    log.debug("continue: attachments not found")
                pass
            pass
        pass

    def init_message_paths(self, connection, log):
        self.file_query = self.file_query % (ZIMBRA_STORE_DIR, self.group_id, self.id)

        with closing(connection.cursor()) as cursor:
            cursor.execute(self.file_query)
            log.info("init_message_paths: rowcount %s query %s" % (str(cursor.rowcount), self.file_query))
            for row in cursor.fetchall():
                self.message_paths.append(row[0])
        pass

    def __init__(self, account_id, group_id, account_email):
        self.message_paths = []
        self.id = account_id
        self.group_id = group_id
        self.email = account_email
        pass
    pass


class ZimbraManager:
    """
    Works with zimbra mysql
    """

    accounts = []
    config = None  # inited ConfigParser
    log = None  # inited logger

    count_query = 'SELECT COUNT(*) FROM zimbra.mailbox'  # count query

    # http://wiki.zimbra.com/wiki/How_to_move_mail_from_one_user's_folder_to_another,_or_to_send_it_for_external_delivery
    base_query = 'SELECT id, group_id, comment FROM zimbra.mailbox'  # base query
    order = 'ORDER BY id'  # base order

    def _init_config_parser(self):
        """
        loads zimbra config
        """
        b = '[' + _tmp_block + '] '
        self.config = ConfigParser()
        self.config.readfp(StringIO(b + Utils.ex_call(ZIMBRA_CONFIG_CMD, logger=self.log)))
        pass

    def _get_socket_file_path(self):
        """
        builds path to mysql socket
        """

        # TODO: need more abstract way
        socket = self.config.get(_tmp_block, _zcp_home)
        socket += '/' + self.config.get(_tmp_block, _zcp_db_dir).split('/', 1)[1]
        socket += '/' + self.config.get(_tmp_block, _zcp_db_socket).split('/', 1)[1]
        return socket

    def _get_mysql_conn(self):
        """
        Init new zimbra mysql connection
        """
        return MySQLdb.connect(
            host=self.config.get(_tmp_block, _zcp_db_host),
            user=self.config.get(_tmp_block, _zcp_db_user),
            passwd=self.config.get(_tmp_block, _zcp_db_pass),
            unix_socket=self._get_socket_file_path(),
            use_unicode=True,
            charset='utf8'
        )

    def _check_user(self):
        user = self.config.get(_tmp_block, _zcp_user)
        if getuser() != user:
            print
            print "ERROR. You MUST HAVE the rights of user '" + user
            print "' Please, do 'su " + user + "' and try again."
            sys.exit(1)

    def process(self):
        with closing(self.connection.cursor()) as cursor:
            cursor.execute(self.base_query)

            self.log.info(
                "process: %s rowcount %s query %s"
                % (self._child, str(cursor.rowcount), self.base_query))

            for row in cursor.fetchall():
                acc = Account(row[0], row[1], row[2])
                self.log.info(
                    "Account: start work with %s id %s group_id %s"
                    % (acc.email, acc.id, acc.group_id))

                acc.init_message_paths(self.connection, self.log)
                acc.delete_attachments(self._ttl, self.log, debug=self._debug)

    def run_self(self):
        if self._limit == _def_limit:
            self._limit = 1000
        self._start = 0

        _start = self._start
        _limit = self._limit
        _count = 0
        _child = 0

        with closing(self.connection.cursor()) as cursor:
            cursor.execute(self.count_query)
            _row = cursor.fetchone()

            if not _row:
                Utils.error_exit('count_query fail', logger=self.log)
            else:
                _count = _row[0]

            self.log.info("accounts count %s" % _count)

            while _count > 0:
                _child += 1
                _cmd = [
                    __file__,
                    "--child=" + str(_child),
                    "--start=" + str(_start),
                    "--limit=" + str(_limit),
                    "--time-to-live=" + str(self._arguments.TIME_TO_LIVE),
                    "--log-dir=" + str(self._log_dir)
                ]

                if self._debug:
                    _cmd.append("--debug=" + str(1))

                _process = subprocess.Popen(_cmd)

                self.log.info(
                    "run_self: child: %s child_pid %s start %s limit %s count %s cmd %s"
                    % (_child, _process.pid, _start, _limit, _count, Utils.array_to_string(src_list=_cmd)))

                _start += _limit
                _count -= _limit
                pass
            pass
        pass

    def _init_processing(self):
        """
        Init processing.

        * accounts defined.
            it's 'hand run' (admin work). work to the end it this process
        * start changed.
            it's 'self run' (self was run this). work to the end in this process
        * all accounts (changed only ttl or/and portion or/and log_file)
            it's 'clear run'. must starts self with change start or/and limit params
        """

        if self._accounts != _def_accounts and self._accounts != _def_all_accounts_flag:
            """
            accounts defined
            """
            if os.path.isfile(self._accounts) and os.access(self._accounts, os.R_OK):
                """
                file exists and readable
                """
                with open(self._accounts) as f:
                    _accounts = Utils.array_to_string(f.read().splitlines())
                    self.base_query = "%s WHERE comment IN (%s) %s;" % (self.base_query, _accounts, self.order)
            else:
                """
                it's single e-mail address
                """
                self.base_query += " WHERE comment LIKE '"

                if self._accounts.endswith(_def_all_accounts_flag):
                    """use search mask"""
                    self.base_query += self._accounts[:-1] + "%';"
                else:
                    self.base_query += self._accounts + "';"
                pass

            self.process()

        elif self._accounts == _def_all_accounts_flag:
            """
            must run self COUNT/LIMIT times
            """
            self.run_self()

        elif self._child != _def_child:
            """
            start and limit changed
            this was run automatically
            """
            self.base_query = "%s %s LIMIT %s, %s;" % (self.base_query, self.order, self._start, self._limit)
            self.process()
        pass

    def __init__(self, arguments):
        self._arguments = arguments
        self._accounts = arguments.ACCOUNTS
        self._start = arguments.START
        self._limit = arguments.LIMIT
        self._ttl = arguments.TIME_TO_LIVE

        self._init_config_parser()
        self._check_user()

        if os.path.isdir(arguments.LOG_DIR):
            self._log_dir = arguments.LOG_DIR
        else:
            self._log_dir = os.path.abspath(arguments.LOG_DIR)
            Utils.ensure_dir(self._log_dir)

        if str(arguments.CHILD) == _def_child:
            self._child = str(0)
        else:
            self._child = str(arguments.CHILD)
            pass
        self._log_file = "%s/%s_%s.log" % (self._log_dir, os.path.basename(__file__), self._child)

        self._debug = (arguments.DEBUG == 1)
        self.log = Utils.get_logger(self._log_file, debug=self._debug)

        self.connection = self._get_mysql_conn()
        self._init_processing()
        pass
    pass
###############################################################################
# </editor-fold>

# <editor-fold desc="Arguments">
###############################################################################

parser = ArgumentParser(
    prog=__program__,
    formatter_class=RawDescriptionHelpFormatter,
    description=dedent("""
    Removes attachments from zimbra email messages

    Usages:

    [zimbra_user@mail_store]$ ./zak.py -a account@my.dom -t 180

        will remove attachments older then 180 days in account@my.dom


    -------------------------

    [zimbra_user@mail_store]$ cat list
    account1@my.dom
    account2@my.dom

    [zimbra_user@mail_store]$ ./zak.py -a list -t 180

        will remove attachments older then 180 days
        in account1@my.dom and in account2@my.dom

    -------------------------

    [zimbra_user@mail_store]$ ./zak.py -a all -t 180

        will remove attachments older then 180 days in all accounts

    -------------------------
    """))
parser.add_argument(
    "-a", "--accounts", dest="ACCOUNTS", type=str, default=_def_accounts,
    help="list zimbra-accounts for processing. default: empty. All Accounts: '"+_def_all_accounts_flag+"'")
parser.add_argument(
    "-t", "--time-to-live", dest="TIME_TO_LIVE", type=int, default=_def_ttl_days,
    help="attachments time to live in days. default: " + str(_def_ttl_days))
parser.add_argument(
    "-s", "--start", dest="START", type=int, default=_def_start,
    help="account list Start position for processing. default: " + str(_def_start))
parser.add_argument(
    "-l", "--limit", dest="LIMIT", type=int, default=_def_limit,
    help="limit of accounts portion for processing. default: " + str(_def_limit))
parser.add_argument(
    "-d", "--log-dir", dest="LOG_DIR", type=str, default=_def_log_dir,
    help="directory for log files. default: " + _def_log_dir)
parser.add_argument(
    "-c", "--child", dest="CHILD", type=str, default=_def_child,
    help="Child counter. Private argument.")
parser.add_argument("--debug", dest="DEBUG", type=int, default=0, help="is debug mode.")
parser.add_argument("-v", "--version", action="version", version="%(prog)s " + __version__)
#  "-h", "--help" auto generated parameter

args = parser.parse_args()

###############################################################################
# </editor-fold>

zm = ZimbraManager(args)
Utils.log = zm.log
Utils.log.info("Successful end of work")
sys.exit(0)