import csv
import smtplib
import configparser
import logging
import logging.config
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import parseaddr
import pandas as pd

### define logger
LOG_FILENAME = datetime.datetime.now().strftime('joyo_%d_%m_%Y.log')
dictLogConfig = {
        "version":1,
        "handlers":{
                    "fileHandler":{
                        "class":"logging.FileHandler",
                        "formatter":"myFormatter",
                        "filename":LOG_FILENAME
                        }
                    },
        "loggers":{
            "email-sender":{
                "handlers":["fileHandler"],
                "level":"INFO",
                }
            },

        "formatters":{
            "myFormatter":{
                "format":"%(asctime)s - %(name)s - %(levelname)s - %(message)s"
                }
            }
        }

logging.config.dictConfig(dictLogConfig)
LOG = logging.getLogger("email-sender")
LOG.info("Program started")
####

CONFIG_FILE = 'config.ini'

EMAIL = {
    'username':'',
    'password':'',
    'port':'',
    'host':'',
    'receipients':'',
    'content':'',
    'template':'',
    'subject':'',
    'attachment':'',
    'copyto':''
}

def read_config():
    try:
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding='utf-8')
        EMAIL["host"] = config['DEFAULT']["SMTP"]
        EMAIL["port"] = config['DEFAULT']["PORT"]
        EMAIL['username'] = config['DEFAULT']["USERNAME"]
        EMAIL["password"] = config['DEFAULT']["PASSWORD"]
        EMAIL["recipients"] = config['DEFAULT']["RECIPIENTS"]
        EMAIL["template"] = config['DEFAULT']["TEMPLATE"]
        EMAIL["subject"] = config['DEFAULT']["SUBJECT"]
        EMAIL["attachment"] = config['DEFAULT']["ATTACHMENT"]
        EMAIL["copyto"] = config['DEFAULT']["COPYTO"]
    except Exception as e:
        LOG.error(f'read config.ini failed, {e}')

def read_template(filename):
    try:
        with open(filename, encoding='utf-8-sig') as template:
            EMAIL["content"] = template.read()
    except Exception as e:
        LOG.error(f'read template file {filename} failed, {e}')

def read_csv(filename):
    try:
        with open(filename, encoding='utf-8-sig') as infile:
            reader = csv.DictReader(infile)
            data = {}
            for row in reader:
                for header, value in row.items():
                    try:
                        data[header].append(value)
                    except KeyError:
                        data[header] = [value]
        return data
    except Exception as e:
        LOG.error(f'read csv file {filename} failed, {e}')

def read_xslx(filename):
    try:
        df = pd.read_excel(filename, encoding='utf-8-sig')
        data = {}
        for header in df.columns:
            for i in df.index:
                try:
                    # print(header + " == " + df[header][i])
                    data[header].append(df[header][i])
                except KeyError:
                    data[header] = [df[header][i]]
        return data
    except Exception as e:
        LOG.error(f'read xslx file {filename} failed, {e}')

def main():
    try:
        read_config()
        if not EMAIL["host"]:
            LOG.error('Cannot find smtp host.')
            return
        
        read_template(EMAIL["template"])
        if not EMAIL["content"]:
            LOG.error('Email body is empty.')
            return

        if EMAIL["recipients"].endswith('.csv'):
            data = read_csv(EMAIL["recipients"])
        else :
            data = read_xslx(EMAIL["recipients"])
        if not 'email' in data:
            LOG.error(f'Cannot find "email" in {EMAIL["recipients"]}')
            return
        
        try:
            if EMAIL["port"]:
                s = smtplib.SMTP(EMAIL["host"], EMAIL["port"])
            else:
                s = smtplib.SMTP(EMAIL["host"])
        except Exception as e:
            LOG.error(f'{EMAIL["host"]}:{EMAIL["port"]} connection failed: {e}')
            return

        try:
            s.login(EMAIL['username'], EMAIL["password"])
        except Exception as e:
            LOG.error(f'login failed: {e}')
            return

        sucessCount = 0
        failureCount = 0
        for i in range(len(data["email"])):
            if not '@' in parseaddr(data["email"][i])[1]: 
                failureCount += 1
                LOG.error(f'Email address is invalid: {data["email"][i]}') 
            else:
                msg = MIMEMultipart()            
                msg['From'] = EMAIL["username"]
                msg['Cc'] = EMAIL["copyto"]
                subject =  EMAIL["subject"]
                content = EMAIL["content"]
                count = 0
                for attachment in EMAIL["attachment"].split(","):
                    if attachment != '':
                        with open(attachment, 'rb') as fp:
                            try:
                                if attachment.endswith(".pdf"):                            
                                    att = MIMEApplication(fp.read(), _subtype="pdf")
                                else:
                                    att = MIMEImage(fp.read())
                                # att.add_header('Content-ID', '<0>')
                                att.add_header('Content-Disposition', 'attachment', filename=attachment)
                                msg.attach(att)
                            except Exception as e:
                                LOG.error(f'attach {attachment} failed, {e}')
                        count += 1
                attachment_count = f'email attachement count {count}'
                msg['To'] = data["email"][i]            
                for key, value in data.items():
                    try:
                        content = content.replace('{'+key+'}', str(value[i]))
                        subject = subject.replace('{'+key+'}', str(value[i]))
                    except Exception as e:
                        LOG.error(f'email body is not completed: {e}')
                
                try:
                    msg['Subject'] = Header(subject, 'utf-8').encode()
                    msg.attach(MIMEText(content, 'plain', 'utf-8'))
                    s.send_message(msg)
                    sucessCount += 1
                    print(f'email sent to {msg["To"]}, {attachment_count}')
                    LOG.info(f'email sent to {msg["To"]}, {attachment_count}')
                except Exception as e:
                    failureCount += 1
                    LOG.error(f'Send email to {msg["To"]} failed, {e}')

        s.quit()
        print(f'{sucessCount} emails sent. {failureCount} emails failed.')
        LOG.info(f'{sucessCount} emails sent. {failureCount} emails failed.')
    except Exception as e: 
        print(e)
        LOG.info(f'Program failed: {e}')
####
main()
LOG.info("Program stopped")