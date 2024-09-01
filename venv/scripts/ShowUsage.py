# usage:
#   cd python_projects/focus_reporting/venv/scripts
#   python3 ShowUsage.py -rpt USAGE -g HOURLY -ds 2024-08-01 -de 2024-08-21 -csv

import os
import sys
import pytz
from pathlib import Path
import argparse
import oci
import csv
from datetime import datetime, timedelta

import sendgrid
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

from msal import ConfidentialClientApplication
import smtplib

import mimetypes
from email import encoders
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

##########################################################################
###                       GLOBAL VARIABLES                             ###
##########################################################################

#region FOCUS reporting parameters #

# define the configuration for the source tenant ( requestor )
g_TenantID_source = "ocid1.tenancy.oc1..aaaaaaaagkhw5j4cxtlzb7c7g5efapzma4wdymyqaiij2ouuarg74cvruc2q"
# define the configuration for the tenant where the data is located
g_TenantID_dest = "ocid1.tenancy.oc1..aaaaaaaadbn4d7ald33gyfgtd7wv5gexwv4lnoftbgntveg2yr4c7dcnf57a"

g_Config_source = {
  'user': 'ocid1.user.oc1..aaaaaaaatwfmlfw57qheletbzriaom4zmhkgfaa2mb5rfzqzr4lbdsle4olq',       #marek.prochyra@aspecta.sk
  'key_file': '~/.oci/keys/aspecta_private_key.pem',                                            #private key in tenant "aspectask"
  'fingerprint': 'c8:3c:0b:4b:92:f1:e6:fb:0c:87:9b:2a:16:87:3a:89',                             #fingerprint in tenant "socialnapoistovna"
  'tenancy': g_TenantID_source,                                                                        #socialnapoistovna
  'region': 'eu-frankfurt-1'
}

g_Config_dest = {
  'user': 'ocid1.user.oc1..aaaaaaaatwfmlfw57qheletbzriaom4zmhkgfaa2mb5rfzqzr4lbdsle4olq',       #marek.prochyra@aspecta.sk
  'key_file': '~/.oci/keys/aspecta_private_key.pem',                                            #private key in tenant "aspectask"
  'fingerprint': 'c8:3c:0b:4b:92:f1:e6:fb:0c:87:9b:2a:16:87:3a:89',                             #fingerprint in tenant "socialnapoistovna"
  'tenancy': g_TenantID_dest,                                                                        #socialnapoistovna
  'region': 'eu-frankfurt-1'
}

g_UsageClient = ""

g_ReportType = "DATE"   # currently not used, originally used for csv files with different set of columns
g_QueryType = "COST"
g_Granularity = "HOURLY"
g_Frequency = "REGULAR"
g_ActualMonth = ""
g_TimeUsageStarted = None
g_TimeUsageEnded = None

g_isCSV = False
g_csvFolder = os.path.join ( os.path.expanduser('~'), "python_projects/focus_reporting/save" )
g_csvFilePath = ""
g_csvFileName = "usage_hourly"
#endregion

#region email-related parameters #
g_SmtpServer = "smtp.office365.com"               # SMTP server
g_SmtpServerOCI = "169.254.169.254"                   # smtp.email.eu-frankfurt-1.oraclecloud.com

g_SmtpPort = 587                                  # for SSL use 465, for TLS use 587

g_SmtpUser = "marek.prochyra@aspecta.sk"          # email address of SMTP user
g_SmtpPassword = "MarkOnIS73"                     # email password of SMTP user or app-specific password

g_SmtpUserSendGrid = "apikey"
g_SmtpPasswordSendGrid = "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"  # email password of SMTP user or app-specific password

g_SmtpUserOCI = "ocid1.user.oc1..aaaaaaaaere3lplxkfvgqndsd5len33gsyninkqhdj25ukbrh5bklwn2nsma@ocid1.tenancy.oc1..aaaaaaaagkhw5j4cxtlzb7c7g5efapzma4wdymyqaiij2ouuarg74cvruc2q.m7.com"
g_SmtpPasswordOCI = "<nkqwh2f&F{L6<7rz$}v"

g_OATH2_ClientCredential = "XRP8Q~iABmI4ubyrltxMhHgzamMZH5iSOl6Cadtq"
g_OATH2_TenantID = "9b449085-b541-4408-be3b-3aa2e0f2f689"
g_OATH2_ClientID = "d0f9bfe9-1c88-45c1-9a95-3e479620da65"  # application ID ( Azure app_name = "SendEmail_Auth_App" )
g_OATH2_Scopes    = ["https://graph.microsoft.com/.default"]  # ["https://outlook.office365.com/.default"]      "offline_access%20Mail.ReadWrite%20Mail.send"
g_OATH2_Authority = f"https://login.microsoftonline.com/{g_OATH2_TenantID}"

g_FromEmailAddress = "marek.prochyra@aspecta.sk" # "no-reply@aspecta.sk"        # sender email address
g_ToEmailAddress = ['marek.prochyra@aspecta.sk', 'notifikaciesocpoist@aspecta.sk']    # recipient email address
#endregion

#region OCI
# PAR ( pre-auth. request for access to "Focus_reporting_bucket")
g_bucket_PAR = "https://fr5sgvqmsncb.objectstorage.eu-frankfurt-1.oci.customer-oci.com/p/V9z2vxEoWHGGySOyRK5fDWXNAtjIzCHkuuhminiy6MQbWxlrBCIeF-j4jNSf5gF-/n/fr5sgvqmsncb/b/Focus_reporting_bucket/o"
#endregion


##########################################################################
###                        FUNCTIONS                                   ###
##########################################################################


##########################################################################
# custom argparse *date* type for user dates
##########################################################################
def valid_date_type ( arg_date_str ):
    try:
        return datetime.strptime ( arg_date_str, "%Y-%m-%d" )
    except ValueError:
        msg = "Given Date ({0}) not valid! Expected format, YYYY-MM-DD!".format ( arg_date_str )
        raise argparse.ArgumentTypeError ( msg )

##########################################################################
# validate inputs: ERROR IF: 
    # 1) incorrect inputs, 
    # 2) WEEKLY processing on Monday = 1st of new month => MONTHLY process running at the same day will cover this

##########################################################################
def fn_IsInputValid ( p_CmdlnArgs ):
    
    l_isInputValid = True

    # set start date as 1st of previous month, end date as current date ( = 1st of current month )
    # note: [start, end ] = [ 1st of prev_month, 1st of curr_month ] 
    if g_Frequency == "REGULAR_MONTH":

      # get last successfully extracted date as start date ( midnight ) for monthly extraction      
      l_last_saved_date = datetime.strptime ( os.environ [ 'LAST_SAVED_DATE' ], '%Y-%m-%d' )
      l_TimeUsageStarted = l_last_saved_date.replace ( hour=0, minute=0, second=0, microsecond=0 )
      
      # set today ( 1st of new month ) as END date at midnight
      l_TimeUsageEnded = datetime.today().replace ( hour=0, minute=0, second=0, microsecond=0 )

      #temp
      #l_today = datetime.strptime ( "2024-09-01", "%Y-%m-%d" )
      #endtemp
      
      # print (f"l_last_saved_date {l_last_saved_date}" )
      # print (f"start: {l_TimeUsageStarted}, end: {l_TimeUsageEnded}" )
      
      # sys.exit()

    # set START date as 1st of current month, END date as current day ( = Monday )
    # note: if today ( = monday ) = 1st of month => process will not run
    elif g_Frequency == "REGULAR_WEEK":
      #region set START / END dates

      # check if today is Monday 1st of month => ERROR
      if ( datetime.now().strftime('%A') == 'Monday' and datetime.now().day == 1 ): 
        return False

      else: # OK => set START / END dates
        # START = 1st of current month at midnight
        # END = today (monday) at midnight
        l_TimeUsageStarted = datetime.today().replace ( day=1, hour=0, minute=0, second=0, microsecond=0 )
        l_TimeUsageEnded = datetime.today().replace ( hour=0, minute=0, second=0, microsecond=0 )
      #endregion

    # [start, end ] = [ p_date_start, p_date_end ]
    elif g_Frequency == "ADHOC":
      #region calculate start, end dates based on input parameters ( + validate input parameters )
      if p_CmdlnArgs.date_start and p_CmdlnArgs.date_start > datetime.now():
        print ( "\n!!! Error, Start date cannot be in the future !!!" )
        return False

      if p_CmdlnArgs.date_start and p_CmdlnArgs.date_end and p_CmdlnArgs.date_start > p_CmdlnArgs.date_end:
        print ( "\n!!! Error, Start date cannot be greater than End date !!!" )
        return False

      if p_CmdlnArgs.date_start:        
        l_TimeUsageStarted = p_CmdlnArgs.date_start

      if p_CmdlnArgs.days:
        l_TimeUsageEnded = l_TimeUsageStarted + datetime.timedelta ( days=p_CmdlnArgs.days )

      elif p_CmdlnArgs.date_end:
        l_TimeUsageEnded = p_CmdlnArgs.date_end

      else:
        l_TimeUsageEnded = l_TimeUsageStarted + datetime.timedelta ( days=1 )
    #endregion


    # print (f"start: {l_TimeUsageStarted}, end: {l_TimeUsageEnded}" )
    # sys.exit()

    # set final start and end dates
    global g_TimeUsageStarted
    g_TimeUsageStarted = l_TimeUsageStarted

    global g_TimeUsageEnded
    g_TimeUsageEnded = l_TimeUsageEnded

    # set actual month in a format 'YYYY_MM'
    global g_ActualMonth
    g_ActualMonth = g_TimeUsageStarted.strftime ('%Y-%m-%d') [:7].replace ( '-', '_' )

    # check for max. days allowed for monhtly / yearly data range
    days = ( g_TimeUsageEnded - g_TimeUsageStarted ).days

    if days > 93 and p_CmdlnArgs.granularity == 'DAILY':
        print ( "\n!!! Error, Max 93 days period allowed for DAILY, input is " + str(days) + " days, !!!" )
        return False

    if days > 366 and p_CmdlnArgs.granularity == 'MONTHLY':
        print ( "\n!!! Error, Max 366 days period allowed for MONTHLY, input is " + str(days) + " days, !!!" )
        return False

    return l_isInputValid

##########################################################################
# create instance of UsageClient object
##########################################################################
def fn_GetUsageClient():
  try:      
    # create a Usage API client        
    global g_UsageClient
    g_UsageClient = oci.usage_api.UsageapiClient ( g_Config_dest )

  except Exception as e:
    raise RuntimeError("\nError in fn_GetUsageClient() function: " + str(e))
  
############################################
# main function for data retrieval
############################################
def fn_GetUsageData ( p_DateTimeStarted, p_DateTimeEnded ):

  # print ( f"g_TimeUsageStarted: {g_TimeUsageStarted.strftime ( '%Y-%m-%dT%H:%M:%SZ' )}, g_TimeUsageEnded: {g_TimeUsageEnded.strftime ( '%Y-%m-%dT%H:%M:%SZ' )}" )
  # print ( f"From: {p_DateTimeStarted.strftime ( '%Y-%m-%dT%H:%M:%SZ' )}, to: {p_DateTimeEnded.strftime ( '%Y-%m-%dT%H:%M:%SZ' )}")

  try:
    # oci.usage_api.models.RequestSummarizedUsagesDetails
    l_RequestSummarizedUsagesDetails = oci.usage_api.models.RequestSummarizedUsagesDetails (
    tenant_id = g_TenantID_dest,
    granularity = g_Granularity,
    query_type = g_QueryType,
    group_by= [ 'resourceId', 'service', 'skuPartNumber', 'skuName' ],
    time_usage_started = p_DateTimeStarted.strftime('%Y-%m-%dT%H:%M:%SZ'),
    time_usage_ended = p_DateTimeEnded.strftime('%Y-%m-%dT%H:%M:%SZ')
    )

    # usageClient.request_summarized_usages
    l_RequestSummarizedUsages = g_UsageClient.request_summarized_usages (
      l_RequestSummarizedUsagesDetails,
      retry_strategy=oci.retry.DEFAULT_RETRY_STRATEGY
    )

  except Exception as e:
    raise RuntimeError ( "\nError in fn_GetUsageData() function: " + str(e) )

  finally: 
    return l_RequestSummarizedUsages

##########################################################################
# get monthly data aggregated as hourly chunk ( <= 36 hours limitation for HOURLY data )
##########################################################################
def fn_GetHourlyAggregMonthlyData ():

    l_HourlyData = []
    l_CurrentStart = g_TimeUsageStarted

    # Process in 36-hour chunks
    while l_CurrentStart < g_TimeUsageEnded:
      l_CurrentEnd = min ( l_CurrentStart + timedelta ( hours=36 ), g_TimeUsageEnded )
      print ( f"Fetching data from {l_CurrentStart} to {l_CurrentEnd}" )
      l_HourlyDataItems_36hrs = fn_GetUsageData ( l_CurrentStart, l_CurrentEnd ).data.items
      l_HourlyData.extend ( l_HourlyDataItems_36hrs )
      l_CurrentStart = l_CurrentEnd
    
    return l_HourlyData

##########################################################################
# main function that retrieves usage/cost data based on input parametres
##########################################################################
def fn_GetData_Main ():

  try:
    if g_Frequency == "REGULAR_MONHTLY" or g_Frequency == "REGULAR_WEEKLY" or g_Granularity == "HOURLY":
      return fn_GetHourlyAggregMonthlyData ()
    else: return fn_GetUsageData ( g_TimeUsageStarted, g_TimeUsageEnded )

  except Exception as e:
    raise RuntimeError ( "\nError in fn_GetData() function: " + str(e) )
  
##########################################################################
# print or save data
##########################################################################
def fn_PrintSaveData ( p_UsageData ):
  # programmatically filter the response by service types
  if g_Granularity == "HOURLY":
    l_FilteredData = [ item for item in p_UsageData if item.sku_name in [ "Database Exadata Cloud at Customer - Database OCPU" ]]
    #print ( p_UsageData )

  else:
    l_FilteredData = [ item for item in p_UsageData.data.items if item.sku_name in [ "Database Exadata Cloud at Customer - Database OCPU" ]]    
    #print ( p_UsageData.data )
  
  l_FilteredData.sort ( key = lambda x: ( x.ad, x.time_usage_started ) )
  #print ( l_FilteredData )

  # define column headers
  l_lst_Headers = [ "environment", "time_usage_started", "time_usage_ended", "resourceId", "availability_domain", "VCN", "service", "sku_name", "sku_part_number", "computed_amount", "computed_quantity", "currency" ]   


  #get virtual cluster network (VCN) detail ( PROD / TEST ) based on its resource_id
  l_database_client = oci.database.DatabaseClient ( g_Config_dest )

  # save data into file
  if g_isCSV:        
        
    # in WEEKLY process data is overwritten, in MONTHLY process the last remaining data until end of month are appended
    l_file_mode = "a" if ( g_Frequency == "REGULAR_MONTH" ) else "w"

    with open ( g_csvFilePath, l_file_mode, newline='' ) as l_file:
      l_writer = csv.writer ( l_file )

      # in WEEKLY process add header as first line
      if ( g_Frequency != "REGULAR_MONTH" ): l_writer.writerow ( l_lst_Headers )

      for l_item in l_FilteredData:

        l_VCN = l_database_client.get_vm_cluster ( l_item.resource_id ).data

        if   "tqvm" in l_VCN.display_name: l_environment = "TEST" 
        elif "pqvm" in l_VCN.display_name: l_environment = "PROD" 
        else: l_environment = ""

        #print ( l_VCN )        

        # directly write each row to CSV
        l_row = [
          l_environment,
          str ( l_item.time_usage_started ),
          str ( l_item.time_usage_ended ),

          l_item.resource_id,
          l_VCN.availability_domain,
          l_VCN.display_name,    
          l_item.service, 
          l_item.sku_name,
          l_item.sku_part_number,
          str ( l_item.computed_amount ),
          str ( l_item.computed_quantity ),
          l_item.currency
        ] 

        l_writer.writerow ( l_row )
        #print ( l_row )

  else: # print to screen

    # Determine the width of each column
    col_widths = [max ( len ( header ) for header in l_lst_Headers )] * len ( l_lst_Headers )

    # format the header
    l_lst_FormattedHeader = " | ".join ( l_header.ljust ( col_widths[i] ) for i, l_header in enumerate ( l_lst_Headers ))
    l_arr_FormattedRows = [ l_lst_FormattedHeader ]
    l_arr_FormattedRows.append ( "-" * len ( l_lst_FormattedHeader ))


    # format each row
    for l_item in l_FilteredData:

      l_VCN = l_database_client.get_vm_cluster ( l_item.resource_id ).data      

      if   "tqvm" in l_VCN.display_name: l_environment = "TEST" 
      elif "pqvm" in l_VCN.display_name: l_environment = "PROD" 
      else: l_environment = ""

      l_row = [
        l_environment,
        str ( l_item.time_usage_started ),
        str ( l_item.time_usage_ended ),
        l_item.resource_id,
        l_VCN.availability_domain,
        l_VCN.display_name,
        l_item.service, 
        l_item.sku_name,
        l_item.sku_part_number,
        str ( l_item.computed_amount ),
        str ( l_item.computed_quantity ),
        l_item.currency          
      ]
      #print ( l_row )

      #l_FormattedRow = " | ".join ( l_cell.ljust ( col_widths[i] ) for i, l_cell in enumerate ( l_row ) )
      l_FormattedRow = " | ".join (
        ( l_cell if l_cell is not None else "" ).ljust ( col_widths[i] )
          for i, l_cell in enumerate ( l_row )
      )

      l_arr_FormattedRows.append ( l_FormattedRow )

    ### display  original data ###
      #print ( p_UsageData )
      #print ( l_FilteredData )

    ### display formatted data ###
    print ( l_arr_FormattedRows )

##########################################################################
# save file to ObjectStorage
##########################################################################
def fn_SaveFileToObjectStorage():
# Save to Object Storage
  object_storage_client = oci.object_storage.ObjectStorageClient(config={})
  namespace = "your_namespace"  # Replace with your OCI Object Storage namespace
  bucket_name = "your_bucket_name"  # Replace with your bucket name
  object_name = "test.csv"

# Convert CSV content to bytes
  csv_bytes = "" # csv_content.encode('utf-8')

# Upload the CSV file to the bucket
  object_storage_client.put_object(
      namespace_name=namespace,
      bucket_name=bucket_name,
      object_name=object_name,
      put_object_body=csv_bytes
  )

##########################################################################
# ObjectStorage manipulations - NOT WORKING
##########################################################################
def fn_ObjectStorageManipulations_test():
  #region test Object Storage manipulations
  '''
  l_config = oci.config.from_file ( profile_name= "DEFAULT" )  # For primary tenant
  #l_config = oci.config.from_file ( profile_name="REMOTE" )   # For remote tenant

  l_object_storage_client = oci.object_storage.ObjectStorageClient ( l_config )

  print ( str(l_object_storage_client) )

  # Get the namespace
  #l_namespace = l_object_storage_client.get_namespace().data
  #print ( l_namespace )

  #sys.exit()

  # List buckets in the specified compartment
  l_namespace = "fr5sgvqmsncb"
  #l_compart_ocid = "ocid1.compartment.oc1..aaaaaaaaygvp3triltu7zfhno5eu5iks7dcj2yop3mtxfd6io7ifwncd43hq"  # compartment: MPR_Comp_VM_Lnx_Private
  l_compart_ocid = "ocid1.tenancy.oc1..aaaaaaaagkhw5j4cxtlzb7c7g5efapzma4wdymyqaiij2ouuarg74cvruc2q"      # compartment: root

  try:
    l_arr_buckets = l_object_storage_client.list_buckets ( l_namespace, l_compart_ocid ) 
    print (l_arr_buckets)

    # Print bucket names
    for l_bucket in l_arr_buckets.data: 
      print(f'Bucket Name: { l_bucket.name}' )

    # Print bucket names
    for l_bucket in l_arr_buckets.data: 
      print(f'Bucket Name: { l_bucket.name}' )

  except oci.exceptions.ServiceError as e:
    print(f'Error: {e.code} - {e.message}')
  '''
  #endregion


##########################################################################
# send email with CSV file as attachment
##########################################################################

def fn_SendMail_NOT_WORKING(): 
#region !!!!!!!!!!!!!! NOT WORKING !!!!!!!!!!!!!! 
  # send email using SMTP - version 1
  def fn_SendEmailSMTP_v1 ():

    global g_csvFilePath
    g_csvFilePath = os.path.join ( g_csvFolder, g_csvFileName )

    #region create the email message
    l_msg = EmailMessage()
    l_msg ['From'] = g_FromEmailAddress
    l_msg ['To'] = g_ToEmailAddress
    l_msg ['Subject'] = f"New SP Focus Report (csv): {os.path.basename ( g_csvFilePath )}"
    l_msg.set_content ( f"A new CSV file has been created: {g_csvFilePath}")

    try:
      # Attach the CSV file
      with open ( g_csvFilePath, 'rb') as l_file:
        l_file_data = l_file.read()
        l_file_name = os.path.basename ( g_csvFilePath )
        l_mime_type, _ = mimetypes.guess_type ( g_csvFilePath )
        l_msg.add_attachment ( l_file_data, maintype = l_mime_type.split ( '/') [0], subtype = l_mime_type.split ( '/' ) [1], filename = l_file_name )

    except Exception as e:
      print ( f"fn_SendEmailWithCsvFile(): Failed to access CSV file: {e}" )
      sys.exit()
    #endregion

    try:
      # Send the email
      with smtplib.SMTP ( g_SmtpServer, g_SmtpPort ) as l_server:
        l_server.starttls()  # secure the connection
        l_server.login ( g_SmtpUser, g_SmtpPassword )
        l_server.send_message ( l_msg )
        l_server.quit()

    except Exception as e:
      print ( f"fn_SendEmailWithCsvFile(): Failed to send email: {e}" )
      sys.exit()

    print(f"Email sent with attachment: {g_csvFilePath}")

  # send email using SMTP - version 2
  def fn_SendEmailSMTP_v2 ():

    # Create message container
    l_msg = MIMEMultipart()
    l_msg ['From'] = g_FromEmailAddress
    l_msg ['To'] = g_ToEmailAddress
    l_msg ['Subject'] = "Test suject"

    l_body = 'This is a test email sent from Python using SMTP credentials.'

    # Attach email body
    l_msg.attach ( MIMEText ( l_body, 'plain' ))

    # Setup the server and send the email
    try:
        l_server = smtplib.SMTP ( g_SmtpServerOCI, g_SmtpPort )
        print ("ok")
        sys.exit()

        l_server.starttls()  # Secure the connection
        l_server.login ( g_SmtpUserOCI, g_SmtpPasswordOCI )

        l_msg_text = l_msg.as_string()
        l_server.sendmail ( g_FromEmailAddress, g_ToEmailAddress, l_msg_text )
        print ( "Email sent successfully." )

    except Exception as e:
      print(f"Error: {e}")

    finally:
        l_server.quit()

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
#endregion

#region WORKING using SendGridApp
def fn_SendMail_SendGridApp():
  # initialize the SendGrid client with your API key
  
  # region define SMTP server details
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
  #endregion
  
  print ( g_csvFilePath )

  # region send the email using SendGrid's SMTP server
  try:
    #open CSV file
    with open ( g_csvFilePath, 'rb') as l_file:
      l_file_data = l_file.read()

      # and create / attach a MIMEApplication object for the CSV file
      part = MIMEApplication ( l_file_data, Name = g_csvFileName )        
      msg.attach ( part )
    
    server = smtplib.SMTP ( smtp_server, smtp_port )      
    server.starttls()  # Upgrade the connection to TLS
    server.login ( smtp_username, smtp_password )
    server.sendmail ( msg['From'], g_ToEmailAddress, msg.as_string() )
    server.quit()
    print("Email sent successfully!")

  except FileNotFoundError:
    print ( f"Error: The file '{g_csvFilePath}' was not found.")

  except Exception as e:
      print ( f"Failed to send email: {e}" )  

#endregion


##########################################################################
# Main Process
##########################################################################
def fn_Main():

  # local function fn_ProcessInputParams()
  #  1) sets START / END date for data extraction
  #  2) returns False ( error ) in case of:
          # incorrect inputs or, 
          # 2) WEEKLY processing on Monday = 1st of new month => MONTHLY process at the same day will cover this
  def fn_ProcessInputParams():
    l_arr_ReportType = ['ALL', 'DATE', 'SERVICE', 'PRODUCT', 'REGION', 'RESOURCE', 'SPECIAL', 'TENANT', 'COMPARTMENT']
    l_arr_Frequency = ['REGULAR', 'ADHOC']  # REGULAR = each monday with HOURLY granularity, ADHOC = anytime
    l_arr_Granularity = ['HOURLY', 'DAILY', 'MONTHLY']

    #region get Command Line Parser
    parser = argparse.ArgumentParser ( usage=argparse.SUPPRESS, formatter_class=lambda prog: argparse.HelpFormatter(prog, max_help_position=80, width=150 ))
    parser.add_argument ( '-c', default="", dest='config_file', help='OCI CLI Config file' )
    parser.add_argument ( '-t', default="", dest='config_profile', help='Config Profile inside the config file' )
    parser.add_argument ( '-p', default="", dest='proxy', help='Set Proxy (i.e. www-proxy-server.com:80) ' )
    parser.add_argument ( '-ip', action='store_true', default=False, dest='is_instance_principals', help='Use Instance Principals for Authentication' )
    parser.add_argument ( '-dt', action='store_true', default=False, dest='is_delegation_token', help='Use Delegation Token for Authentication' )

    parser.add_argument ( "-report", default="ALL", dest='report', help="Report Type = " + ' / '.join(x for x in l_arr_ReportType) + " ( Default = ALL )" )
    parser.add_argument ( "-qrt", default="USAGE", dest='query_type', help="USAGE, COST (Default USAGE)" )
    parser.add_argument ( "-g", default="DAILY", dest='granularity', help="Granularity HOURLY, DAILY or MONTHLY (Default DAILY)" )
    parser.add_argument ('-freq', default='REGULAR', dest='frequency', help="Report frequency REGULAR or ADHOC (Default REGULAR)")
     
    parser.add_argument ( "-ds", default=None, dest='date_start', help="Start Date - format YYYY-MM-DD", type=valid_date_type )
    parser.add_argument ( "-de", default=None, dest='date_end', help="End Date - format YYYY-MM-DD, (Not Inclusive)", type=valid_date_type )
    parser.add_argument ( "-days", default=None, dest='days', help="Add Days Combined with Start Date (de is ignored if specified)", type=int )
    parser.add_argument ( '-csv', action='store_true', default=False, dest='csv', help='Write to CSV files instead of output to the screen - usage_*.csv' )
    cmdln_args = parser.parse_args()
    #endregion

    #region set global variables
    global g_ReportType
    g_ReportType = cmdln_args.report   # currently not used

    global g_QueryType
    g_QueryType = cmdln_args.query_type
    
    global g_Frequency
    g_Frequency = cmdln_args.frequency

    # regular frequency = each monday with hourly granularity
    global g_Granularity    
    if g_Frequency == "REGULAR_MONTH" or g_Frequency == "REGULAR_WEEK": g_Granularity = "HOURLY"
    else: g_Granularity = cmdln_args.granularity

    global g_isCSV
    g_isCSV = cmdln_args.csv
    #endregion

    # validate input parameters
    l_all_ok = fn_IsInputValid ( cmdln_args )

    print (l_all_ok)
    return l_all_ok
  
  #################################
  # main part of function fn_Main()
  #################################
  
  # if input parameters are OK => START / END dates of extraction are set
  if fn_ProcessInputParams() == False: sys.exit()
    
  # all is OK => continue

  # set CSV file path     
  global g_csvFileName
  g_csvFileName = ( g_QueryType + "_" + g_Granularity ).lower() + "_" + g_ActualMonth + ".csv"
  
  global g_csvFilePath
  g_csvFilePath = os.path.join ( g_csvFolder, g_csvFileName )

  ### I. start extraction process => get USAGE client
  fn_GetUsageClient()

  print ( f"Extracting data for period: {g_TimeUsageStarted} - {g_TimeUsageEnded}" )

  ### II. get data based on input parameters
  l_UsageData = fn_GetData_Main ()  
  print ( "Usage data for requested days is extracted" )

  ### III.print or save data based on input parameters
  fn_PrintSaveData ( l_UsageData )  

  # store WEEKLY_PROCESS END DATE into system variable for next extraction
  # Note: since extraction end date takes D-1 data, this variable will be used as START DATE for monthly extraction process
  os.environ['LAST_SAVED_DATE'] = datetime.now().strftime('%Y-%m-%d')


  ### IV. send email on 1st of each month
  if datetime.now().day == 1: fn_SendMail_SendGridApp ()


##########################################################################
# Main Process
##########################################################################  

os.environ['LAST_SAVED_DATE'] = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
print (os.environ['LAST_SAVED_DATE'])

fn_Main()
print ( "Process has finished successfully" )
