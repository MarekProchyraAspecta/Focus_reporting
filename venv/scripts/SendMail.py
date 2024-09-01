import os
import sys

import sendgrid
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

from msal import ConfidentialClientApplication
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


#region SMTP server parameters #
g_SmtpServer = "smtp.office365.com"               # SMTP server
g_SmtpServerOCI = "169.254.169.254"                   # smtp.email.eu-frankfurt-1.oraclecloud.com
g_SmtpServerSendGrid = "smtp.sendgrid.net"

g_SmtpPort = 587                                  # for SSL use 465, for TLS use 587

g_SmtpUser = "marek.prochyra@aspecta.sk"          # email address of SMTP user
g_SmtpUserSendGrid = "apikey"
g_SmtpPasswordSendGrid = "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"  # email password of SMTP user or app-specific password

g_OATH2_ClientCredential = "XRP8Q~iABmI4ubyrltxMhHgzamMZH5iSOl6Cadtq"
g_OATH2_TenantID = "9b449085-b541-4408-be3b-3aa2e0f2f689"
g_OATH2_ClientID = "d0f9bfe9-1c88-45c1-9a95-3e479620da65"  # application ID ( Azure app_name = "SendEmail_Auth_App" )
g_OATH2_Scopes    = ["https://graph.microsoft.com/.default"]  # ["https://outlook.office365.com/.default"]      "offline_access%20Mail.ReadWrite%20Mail.send"
g_OATH2_Authority = f"https://login.microsoftonline.com/{g_OATH2_TenantID}"

g_SmtpUserOCI = "ocid1.user.oc1..aaaaaaaaere3lplxkfvgqndsd5len33gsyninkqhdj25ukbrh5bklwn2nsma@ocid1.tenancy.oc1..aaaaaaaagkhw5j4cxtlzb7c7g5efapzma4wdymyqaiij2ouuarg74cvruc2q.m7.com"
g_SmtpPasswordOCI = "<nkqwh2f&F{L6<7rz$}v"

g_FromEmailAddress = "marek.prochyra@aspecta.sk"        # sender email address
g_ToEmailAddress = ['marek.prochyra@aspecta.sk', 'notifikaciesocpoist@aspecta.sk']    # recipient email address
#endregion

g_csvFolder = os.path.join ( os.path.expanduser('~'), "python_projects/focus_reporting/save" )
g_csvFilePath = os.path.join ( g_csvFolder, "usage_daily_2024_05.csv" )

# send email using Azure access token
def fn_SendMail_AzureToken ():

  # local function fn_GetOath2Token(): get access token
  def fn_GetOath2Token():    
    try:
      l_app = ConfidentialClientApplication (
        g_OATH2_ClientID,
        authority = g_OATH2_Authority,
        client_credential = g_OATH2_ClientCredential
      )
      print ( "App was created" )

      try: 
        # Acquiring token for a specific account
        l_token_response = l_app.acquire_token_for_client ( scopes = g_OATH2_Scopes )
 
        if 'access_token' in l_token_response: 
          print ( "Token acquired successfully" ) 
          return  l_token_response ['access_token']        
        else: 
          print("Error Description:", l_token_response.get ( 'error_description' ))
          print("Error:", l_token_response.get ( 'error' ))

      except Exception as e: 
        print ( f"Failed to obtain access token: {e}" )

    except Exception as e: 
      print ( f"Failed to create app: {e}" )
        
  ### MAIN FUNCTION    
  
  #region create the email message
  l_msg = MIMEMultipart()
  l_msg ['From'] = g_SmtpUser
  l_msg ['To'] = g_ToEmailAddress
  l_msg ['Subject'] = f"New SP Focus Report (csv): {os.path.basename ( g_csvFilePath )}"

  #print ( g_csvFilePath)

  try:  # Attach the CSV file
    with open ( g_csvFilePath, 'rb') as l_file:
        # create a MIMEBase object
        l_part = MIMEBase ( 'application', 'octet-stream' )
        l_part.set_payload ( l_file.read() )

        # Encode the payload using Base64
        encoders.encode_base64 ( l_part )

        # Add header for the attachment
        l_part.add_header(
            'Content-Disposition',
            f'attachment; filename= {g_csvFilePath.split("/")[-1]}',
        )

        # Attach the file to the email message
        l_msg.attach(l_part)

        # print ( l_msg )    

  except Exception as e:
    print ( f"fn_SendEmailWithCsvFile(): Failed to access CSV file: {e}" )    
  #endregion

  try: # send email
    with smtplib.SMTP ( g_SmtpServer, g_SmtpPort ) as l_server:      
      l_server.connect ( g_SmtpServer, g_SmtpPort )       
      l_server.ehlo()
      l_server.starttls()
      l_server.ehlo()

      print ( f"SMTP server created: {l_server}" )

      # Generate OAuth2 authentication string
      l_access_token = fn_GetOath2Token()
      #print ( l_token )  
      l_auth_string = f"user={g_SmtpUser}\1auth=Bearer {l_access_token}\1\1"

      #print ( l_auth_string + "\n" )      
      
      try:  # authenticate with the OAuth2 token
        l_response = l_server.docmd ( "AUTH XOAUTH2", l_auth_string.encode('utf-8') )
        print ( f"DoCmd function: {l_response}" + "\n" )

        # Send the email
        print ( "Sending email ...")

        try:          
          l_server.sendmail ( g_SmtpUser, g_ToEmailAddress, l_msg.as_string())
          print ( f"Email sent with attachment: {g_csvFilePath}" )

        except Exception as e: print ( f"Failed to send email: {e}" )          

      except Exception as e: print ( f"Failed to execute docmd(): {e}" )

      finally:            
        l_server.quit()        

  except smtplib.SMTPAuthenticationError as e:
    print(f"Failed to login: {e.smtp_code} - {e.smtp_error.decode()}")

  except smtplib.SMTPSenderRefused as e:
    print(f"Sender refused: {e.smtp_code} - {e.smtp_error.decode()}")

  except Exception as e:
    print ( f"fn_SendEmailWithCsvFile(): Failed to send email: {e}" )
    sys.exit()

# send email using SendGrid app
def fn_SendMail_SendGridApp():
  # initialize the SendGrid client with your API key
  
  #region
  # define SMTP server details
  smtp_server = 'smtp.sendgrid.net'
  smtp_port = 587  # Use 465 for SSL or 587 for TLS
  smtp_username = 'apikey'  # This is literally 'apikey'
  smtp_password = g_SmtpPasswordSendGrid

  # Set up the email content
  msg = MIMEMultipart()
  msg['From'] = g_FromEmailAddress    # "no-reply@aspecta.sk"   !!! create this user in Aspecta Azure portal ( in AD )
  msg['To'] = ', '.join ( g_ToEmailAddress ) # g_ToEmailAddress
  msg['Subject'] = 'SP - Monthly Focus report'
  body = 'Please find the attached CSV file.'
  msg.attach ( MIMEText ( body, 'plain' ))

  print (g_csvFilePath)
  try:
    with open ( g_csvFilePath, 'rb') as l_file:
      l_file_data = l_file.read()
  except Exception as e:
    print(f"Failed to open file: {e}")  

  # Create a MIMEApplication object for the CSV file
  part = MIMEApplication ( l_file_data, Name = 'usage_daily_2024-05.csv' )
  
  # Send the email using SendGrid's SMTP server
  try:
      msg.attach ( part )
      
      server = smtplib.SMTP ( smtp_server, smtp_port )      
      server.starttls()  # Upgrade the connection to TLS
      server.login ( smtp_username, smtp_password )
      server.sendmail ( msg['From'], g_ToEmailAddress, msg.as_string() )
      server.quit()
      print("Email sent successfully!")
  except Exception as e:
      print(f"Failed to send email: {e}")  
  #endregion

###### MAIN PROCESS ######
fn_SendMail_SendGridApp ()
###### MAIN PROCESS ######
