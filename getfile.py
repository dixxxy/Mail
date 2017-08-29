#!/usr/bin/env python3
# -*- coding: GB18030 -*-

import poplib
import email
import os
import configparser


from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    # �ȴ�msg�����ȡ����:
    charset = msg.get_charset()
    if charset is None:
        # �����ȡ�������ٴ�Content-Type�ֶλ�ȡ:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def get_email_headers(msg):
    # �ʼ���From, To, Subject�����ڸ�������:
    headers = {}
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['date'] = value
            if header == 'Subject':
                # ��Ҫ����Subject�ַ���:
                subject = decode_str(value)
                headers['subject'] = subject
            else:
                # ��Ҫ����Email��ַ:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
                if header == 'From':
                    from_address = value
                    headers['from'] = from_address
                else:
                    to_address = value
                    headers['to'] = to_address
    content_type = msg.get_content_type()
    print('head content_type:%s '% (content_type))
    return headers

def category(msg):
    new_headers = get_email_headers(msg)
    new_from = []
    new_from = new_headers.get('from')
    return new_from
# indent����������ʾ:
def get_email_content(message, base_save_path):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        #message_headers = get_email_headers(message)
        j = j + 1
        file_name = part.get_filename()
        print(file_name)
        charset = part.get_charset()
        contentType = part.get_content_type()
        print('-----���渽��-----')
        if file_name:  # Attachment
            # Decode filename
            h = email.header.Header(file_name)
            print(h)
            dh = email.header.decode_header(h)
            print(dh)
            filename = dh[0][0]
            if dh[0][1]:  # �����������ĸ�ʽ�����ոø�ʽ����
                filename = str(filename, dh[0][1])
            data = part.get_payload(decode=True)

            print('�ļ�����%s' % (filename))
            print("��ǰ����Ŀ¼ : %s" % os.getcwd())
            os.chdir(base_save_path)
            print("��ǰ����Ŀ¼ : %s" % os.getcwd())
            print('�ļ������ͣ�%s' % type(filename))
            try:
                att_file = open(filename,'wb')
                attachment_files.append(filename)
            except:
                print('���������зǷ��ַ����Զ���һ��')
                filename = decode_str(filename)
                print('�ļ�����%s' % (filename))
                print('�ļ������ͣ�%s' % type(filename))
                att_file = open(filename,'wb')
            try:
                att_file.write(data)
            except:
                print('�ļ�����������')
                data = b'part.get_payload(decode=True)'
                att_file.write(data)

            att_file.close()


        elif contentType == 'text/plain' or contentType == 'text/html':
            # ��������
            data = part.get_payload(decode=True)
            charset = guess_charset(part)
            if charset:
                charset = charset.strip().split(';')[0]
                print('charset:%s' % (charset))
                data = data.decode(charset, 'ignore')
            content = data
    return content, attachment_files


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('Users.ini')

    config.write(open('Users.ini','w'))

    emailaddress = config.get('CNNIC2','user') #��д�����ַ
    print(emailaddress)
    password = config.get('CNNIC2','password') #��д��������
    print(password)
    pop3_server = config.get('CNNIC2','pop')   #���������
    print(pop3_server)

    # ���ӵ�POP3������:
    #server = poplib.POP3_SSL(pop3_server,port=995)      #���Ҫ����QQ����ʱ����Ҫ���
    server = poplib.POP3(pop3_server)
    # ���Դ򿪻�رյ�����Ϣ:
    # server.set_debuglevel(1)
    # POP3�������Ļ�ӭ����:
    print(server.getwelcome())
    # �����֤:
    server.user(emailaddress)
    server.pass_(password)
    # stat()�����ʼ�������ռ�ÿռ�:
    messagesCount, messagesSize = server.stat()
    print('messagesCount:%s' %  (messagesCount))
    print('messagesSize:%s'% (messagesSize))
    # list()���������ʼ��ı��:
    resp, mails, octets = server.list()
    print('------ resp ------')
    print(resp)  # +OK 46 964346 ��Ӧ��״̬ �ʼ����� �ʼ�ռ�õĿռ��С
    print('------ mails ------')
    print(mails)  # �����ʼ��ı�ż���С�ı��list��['1 2211', '2 29908', ...]
    print('------ octets ------')
    print(octets)

    new_headers_from = []
    # ��ȡ����һ���ʼ�, ע�������Ŵ�1��ʼ:
    length = len(mails)
    for i in range(length):
        resp, lines, octets = server.retr(i + 1)
        # lines�洢���ʼ���ԭʼ�ı���ÿһ��,
        # ���Ի�������ʼ���ԭʼ�ı�:
        msg_content = b"\r\n".join(lines).decode('latin-1')
        # ���ʼ����ݽ���ΪMessage����
        msg = Parser().parsestr(msg_content)

        # �������Message�����������һ��MIMEMultipart���󣬼�����Ƕ�׵�����MIMEBase����
        # Ƕ�׿��ܻ���ֹһ�㡣��������Ҫ�ݹ�ش�ӡ��Message����Ĳ�νṹ��
        print('---------- ����֮�� ----------')
        msg_headers = get_email_headers(msg)
        new_from = category(msg)
        new_headers_from.append(new_from)

        print('subject:%s'% (msg_headers.get('subject')))
        print('from_address:%s'% (msg_headers.get('from')))
        print('to_address:%s'% (msg_headers.get('to')))
        print('date:%s'% (msg_headers.get('date')))

        title = new_headers_from[i]
        print(title)
        #cur_dir = r'G:/GetFile/cnnic/'         #�½��渽�����ļ��� ʹ�þ���·��
        cur_dir = config.get('CNNIC2','dir')
        new_title = title.split('<')
        n_title = new_title[0]

        if n_title.strip() == '':
            p_title = new_title[1].split('@')
            final_title = p_title[0]
        else:
            final_title = n_title
        print(final_title)
        new_path = os.path.join(cur_dir, final_title)
        if not os.path.isdir(new_path):
            os.makedirs(new_path)
        base_save_path = new_path
        print(base_save_path)
        content, attachment_files = get_email_content(msg, base_save_path)
        print('content:%s'% (content))
        print('attachment_files:%s ' %  (attachment_files))



    # �ر�����:
    server.quit()
