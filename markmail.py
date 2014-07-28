#!/usr/bin/python

import argparse
import email, email.encoders, email.mime, email.mime.text, email.mime.base
import getpass
import mimetypes
import os
import re
import shutil
import smtplib
import string
import sys

import xlrd

def main():
	parser = argparse.ArgumentParser(description='Automatic Mark Emailer')
	
	parser.add_argument('username', help='SMTP username')
	parser.add_argument('domain', help='To email domain')
	parser.add_argument('message', help='Text file containing mail body')
	parser.add_argument('id', help='Cell address for student ID')
	parser.add_argument('name', help='Cell address for student name')
	parser.add_argument('mark', help='Cell address for mark')
	parser.add_argument('sheets', help='Marking sheet folder')
	
	parser.add_argument('-s', '--server', help='SMTP server hostname', dest='server', default='localhost')
	parser.add_argument('-p', '--port', help='SMTP server port, 0 = automatic', type=int, dest='port', default=0)
	
	parser.add_argument('--subject', help='Email subject', dest='subject', default='Assessment result')
	parser.add_argument('--ssl', help='Enable SSL connection', dest='ssl', action='store_true', default=False)
	parser.add_argument('-v', help='Verbose output', dest='verbose', action='store_true', default=False)
	
	args = parser.parse_args()
	
	# Read cell addresses
	r = re.compile("([a-zA-Z]+)([0-9]+)")
	m_id = r.match(args.id)
	m_name = r.match(args.name)
	m_mark = r.match(args.mark)
	
	if not m_id or not m_name or not m_mark:
		print 'Invalid cell address for ID or mark'
		return
	
	col_id = string.lowercase.index(m_id.group(1).lower())
	col_name = string.lowercase.index(m_name.group(1).lower())
	col_mark = string.lowercase.index(m_mark.group(1).lower())
	row_id = int(m_id.group(2)) - 1
	row_name = int(m_name.group(2)) - 1
	row_mark = int(m_mark.group(2)) - 1
	
	# Read message file
	if not os.path.isfile(args.message):
		print 'Message body is not a file or does not exist'
		return
	
	with open(args.message, 'r') as f:
		msg = f.read()
	
	# Replace line endings with SMTP compatible ones
	msg = msg.replace('\r\n','\n').replace('\r','\n').replace('\n','\r\n')
	
	# Read marking sheets
	if not os.path.isdir(args.sheets):
		print 'Marking sheet path is not a directory'
		return
	
	sheets = []
	
	for f in (f for f in os.listdir(args.sheets) if os.path.isfile(os.path.join(args.sheets, f))):
		print "Reading %s" % (f)
		path = os.path.join(args.sheets, f)
		
		# Open Excel sheet and extract information
		with xlrd.open_workbook(filename=path) as w:
			s = w.sheet_by_index(0);
			
			try:
				id = s.cell(row_id, col_id).value;
				name = s.cell(row_name, col_name).value;
				mark = s.cell(row_mark, col_mark).value;
				
				if not id:
					print "File %s has empty ID" % (f)
					continue
			except:
				print "Failed to read cells from %s" % (f)
				continue
		
		with open(path, 'rb') as f:
			content = f.read()
	
		sheets.append({
			'file': path,
			'content': content,
			'name': name,
			'email': id + '@' + args.domain,
			'mark': str(mark)
		})
	
	# Select appropriate SMTP port
	if args.port == 0:
		if args.ssl:
			smtpPort = 465
		else:
			smtpPort = 25
	else:
		smtpPort = args.port
	
	smtpString = "%s:%d" % (args.server, smtpPort)
	print 'Connect to: ' + smtpString
	print 'Marking sheets: ' + str(len(sheets))
	
	# Read SMTP password
	smtpPass = getpass.getpass('SMTP password: ')
	
	# Connect to SMTP server
	if args.ssl:
		server = smtplib.SMTP_SSL(smtpString)
	else:
		server = smtplib.SMTP(smtpString)
	
	server.ehlo()
	if not args.ssl:
		# Might be broken (just use SSL for gmail)
		server.starttls()
		server.ehlo()
	
	server.login(args.username, smtpPass)
	
	try:
		donePath = os.path.join(args.sheets, 'done')
		os.mkdir(donePath)
	except OSError:
		pass
	
	for s in sheets:
		# Generate email content
		emailMsg = email.mime.Multipart.MIMEMultipart('alternative')
		emailMsg['Subject'] = args.subject
		emailMsg['From'] = args.username
		emailMsg['To'] = s['email']
		
		# Attach message
		emailMsg.attach(email.mime.Text.MIMEText(msg % s))
		
		# Attach marking sheet file
		attachment = email.mime.base.MIMEBase('application', 'vnd.ms-excel')
		attachment.set_payload(s['content'])
		email.encoders.encode_base64(attachment)
		attachment.add_header('Content-Disposition', "attachment;filename=%s" % (os.path.basename(s['file'])))
		emailMsg.attach(attachment)
		
		# Send email with attachment
		server.sendmail(args.username, [s['email']], emailMsg.as_string())
		
		print 'Sent: ' + s['email']
		
		# Move finished file to done directory
		#print "%s -> %s" % (s['file'], os.path.join(donePath, os.path.basename(s['file'])))
		shutil.move(s['file'], os.path.join(donePath, os.path.basename(s['file'])))
	
	server.close()

if __name__ == '__main__':
  main()
  