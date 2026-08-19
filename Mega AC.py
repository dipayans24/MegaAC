from datetime import date, datetime, timedelta
import glob
import io
import os
import re
import tempfile
import warnings

import gdown
import gspread
import pandas as pd
import streamlit as st
from stqdm import stqdm
import zipfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from gspread_dataframe import get_as_dataframe, set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

warnings.filterwarnings('ignore')


SCOPES = ['https://www.googleapis.com/auth/drive']

def generate_code(Funnel, WSDate):

    PreviousDate = (datetime.strptime(WSDate , "%Y-%m-%d") - timedelta(days=9)).strftime("%Y-%m-%d")
    if Funnel == "AI for Techies Checkout":
        Amount = "BETWEEN 100 AND 950"
    elif Funnel == "BE10X Checkout":
        Amount = "BETWEEN 10 AND 232"
    elif Funnel == "Office Master Checkout":
        Amount = "BETWEEN 100 AND 590"

    code = f"""
        SELECT DISTINCT
        convert(`Wp Wc Orders - Order`.`date_created_gmt`, datetime) AS `CreatedAt` , 
        `Wp Wc Order Addresses - Order`.`first_name` AS `Customer Name`,
        `Wp Wc Order Addresses - Order`.`email` AS `Email`,
        `Wp Wc Order Addresses - Order`.`phone` AS `Phone Number`,
            `Wp Wc Orders - Order`.`total_amount` AS `Amount`,
        `wp_wc_orders_meta`.`meta_value` AS `Age Group`,
            'techies checkout' as "Payment Slug",
        `Wp Wc Orders - Order`.`status` AS `Status`, 
        'techies checkout' as  "Payment Funnel", 
        'Yes' as "Abandon Cart"
        FROM
        `wp_wc_orders_meta`
        LEFT JOIN `wp_wc_order_addresses` AS `Wp Wc Order Addresses - Order` ON `wp_wc_orders_meta`.`order_id` = `Wp Wc Order Addresses - Order`.`order_id`
        LEFT JOIN `wp_wc_orders` AS `Wp Wc Orders - Order` ON `wp_wc_orders_meta`.`order_id` = `Wp Wc Orders - Order`.`id`
        WHERE
        (`wp_wc_orders_meta`.`meta_key` = 'billing_age')
        AND (
            `Wp Wc Orders - Order`.`date_created_gmt`  >= '{PreviousDate}'  
        /* between '2026-05-23' and '2026-05-30' */
        )
        AND `Wp Wc Orders - Order`.`total_amount` {Amount}
        AND `Wp Wc Orders - Order`.`status` NOT IN ('wc-completed')
        order by 'Amount' desc, 'CreatedAt' asc;
        """

    return code
    
def save_upload(fileupload, fileType = None):
    temp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(temp_dir, fileupload.name)

    with open(tmp_path, "wb") as f:
        f.write(fileupload.getvalue())
   
    return tmp_path

def next_sunday():
    """
    Returns the date of the upcoming Sunday in 'YYYY-MM-DD' format.
    If today is already a Sunday, returns today's date.
    """
    today = date.today()
    # Monday=0 ... Sunday=6
    days_until_sunday = (6 - today.weekday()) % 7
    result = today + timedelta(days=days_until_sunday)
    return result.strftime("%Y-%m-%d")

def getGdriveService(GdriveCredentials, delegated_user=None):
    # Authenticates with Google Drive using a service account file
    # Pass delegated_user="someone@yourdomain.com" to impersonate a real user (needed if
    # uploading/downloading against a personal My Drive folder rather than a Shared Drive)

    creds = service_account.Credentials.from_service_account_file(GdriveCredentials, scopes=SCOPES)

    if delegated_user:
        creds = creds.with_subject(delegated_user)

    return build('drive', 'v3', credentials=creds)

def getFilesList(parent_folder_id, service):
    # Retrieves ALL files/folders within a parent folder (paginated, Shared-Drive aware)
    file_list = []
    page_token = None
    while True:
        results = service.files().list(
            q=f"'{parent_folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora='allDrives'
        ).execute()
        file_list.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break
    return file_list

def getSubfolderId(parent_folder_id, folder_name, service):
    # Looks up a named subfolder's ID within a parent folder
    for item in getFilesList(parent_folder_id, service):
        if item['name'] == folder_name:
            return item['id']
    return None

# ---------- Download ----------

def download_file(service, file_id, file_name, clear):
    # Downloads a single file straight to disk, with a Streamlit progress bar
    if file_name not in st.session_state or clear:
        
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

        progress_bar = st.progress(0, text=f"Downloading {file_name}...")

        with open(file_name, 'wb') as f:
            downloader = MediaIoBaseDownload(fd=f, request=request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    pct = int(status.progress() * 100)
                    progress_bar.progress(pct, text=f"Downloading {file_name}... {pct}%")

        progress_bar.progress(100, text=f"{file_name} downloaded")

        st.session_state[file_name] = file_name
        return file_name

    else:
        return st.session_state[file_name]

def getFilefromGdrive(folder_id, service, ProcessParameter, clear):
    # Downloads all files from a named subfolder within folder_id
    subfolder_id = getSubfolderId(folder_id, ProcessParameter, service)
    file_list = getFilesList(subfolder_id, service)

    filePaths = []
    for f in  file_list :
        download_file(service, f['id'], f['name'], clear)
        filePaths.append(f['name'])

    return filePaths, service


def getSheet(sheet_id, sheet_name, credential_Upload):
  scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
  creds = ServiceAccountCredentials.from_json_keyfile_name(credential_Upload, scope)
  client = gspread.authorize(creds)

  try:
        workbook = client.open_by_key(sheet_id)
        values = workbook.worksheet(sheet_name).get_all_values()
        records = workbook.worksheet(sheet_name).get_all_records()
        # sheet_name = datetime.now().strftime("%b-%Y")

  # Read the downloaded XLSX file into a pandas DataFrame
  
        paymentSlugs = pd.DataFrame(values[1:], columns=values[0])
        return paymentSlugs
  except:
      return None

#Pre-Processing Payment Report and getDates Functions
def generatePaymentReport(filePath, paymentSlugs):
  paymentReport = pd.read_csv(filePath[0], sep=",", date_format="%Y-%m-%d %H:%M:%S", dayfirst=True,  low_memory=False)

  #Filter the rows
  paymentReport = paymentReport[(paymentReport["Tags"].fillna("empty").str.contains("l1") | paymentReport["Tags"].isna() )]

  #Changing the data type of phone number column
  paymentReport["Phone Number"] = paymentReport["Phone Number"].astype(str).str.replace(r"\D", "", regex=True)#.apply(lambda x: getPaymentPhoneNumber(x))

  #& (paymentReport.Status.isin(['captured', 'failed']) <---

  #Filtering the records according to the condition -- captured and PaymentAmountThreshold
  paymentReport = paymentReport[(paymentReport.Amount.between(0, 950)) ]
  paymentReport = paymentReport[~(paymentReport.Status.str.contains('refund'))].reset_index(drop=True)

  #Formatting the column as date time
  paymentReport["CreatedAt"] = pd.to_datetime(paymentReport["CreatedAt"], format = "%Y-%m-%d %H:%M:%S", exact=True, dayfirst=True, yearfirst=False)

  st.write(f"After Basic Filtering- {len(paymentReport)}")
  paymentSlugs["PaymentFunnel"] = paymentSlugs["PaymentFunnel"].apply(lambda x: pd.NA if x  == '' else x)
  paymentSlugs.dropna(subset = ["PaymentFunnel"], inplace = True, how="any")

  #(paymentSlugs["isExotic"] == "No") <---
  mergedDf = paymentReport.merge(paymentSlugs.loc[:, ["PaymentFunnel", "Payment Slug"]],
                                      left_on="Payment Slug", right_on="Payment Slug", how="left")

  mergedDf = mergedDf[(mergedDf["PaymentFunnel"].isin(["10xTechies", "AI", "Excel", "Python"]))]

  st.write(f"After Merging payment slugs - {len(mergedDf)}")

  columns = ["CreatedAt", "Customer Name", "Email", "Phone Number", "Amount", "Age Group", "Payment Slug", "Status", "PaymentFunnel"]

  return mergedDf[columns]

def getDates(Funnel, BatchDate , WSDate, ACBatchStartTime, ACBatchEndTime):

  BatchDate[["Date", 'StartDate', 'EndDate']] = BatchDate[["Date", 'StartDate', 'EndDate']]#.astype('M8[s]')
  ExcludedTimings[["Date", 'StartDate', 'EndDate']] = ExcludedTimings[["Date", 'StartDate', 'EndDate']]#.astype('M8[s]')

  FilteredBatchDate = BatchDate[(BatchDate["Date"] == WSDate) & (BatchDate["Funnel"]  == Funnel) ]

  if len(FilteredBatchDate)>0:
    startTime = ACBatchStartTime  # change as needed
    endTime = ACBatchEndTime     # change as needed

    startDate = ACBatchStartTime #FilteredBatchDate["StartDate"].astype('M8[s]').iloc[0].replace(hour=startTime.hour, minute=startTime.minute, second=0)
    endDate = ACBatchEndTime #FilteredBatchDate["EndDate"].astype('M8[s]').iloc[0].replace(hour=endTime.hour, minute=endTime.minute, second=0)

    st.write(f"startDate = {startDate} and endDate = {endDate}")
  else:
    startDate = None
    endDate = None

  FilteredExcludedTimings = ExcludedTimings[(ExcludedTimings["Date"] == WSDate) & (ExcludedTimings["Funnel"]  == Funnel)]

  if len(FilteredExcludedTimings) > 0:
    excludedStartDates = FilteredExcludedTimings["StartDate"]
    excludedEndDates = FilteredExcludedTimings["EndDate"]
  else:
    excludedStartDates = None
    excludedEndDates = None

  return startDate, endDate, excludedStartDates, excludedEndDates

def CountIf(Main_File, Current_File, MFCol, CFCol, filename):

    MainFileColCleaned = MFCol.replace(" ", "") # Clean whitespace from main column name
    CurrentFileColCleaned = CFCol.replace(" ", "") # Clean whitespace from current column name

    NewColName = MainFileColCleaned[:5]+"_"+re.sub(r"\W", "", filename[:5])+"_"+CurrentFileColCleaned # Construct unique column name

    newcol = 1 # Initialize suffix counter
    while NewColName in Main_File.columns:
        NewColName = NewColName+"_"+str(newcol) # Append suffix if name exists
        newcol = newcol+1 # Increment counter

    lookup_set = set(Current_File[CFCol].astype(str).str.lower().str.strip()) # Create optimized lookup set
    Main_File.insert(loc=len(Main_File.columns), column=NewColName,
                     value=Main_File[MFCol].astype(str).str.lower().str.strip().isin(lookup_set).astype(int),
                     allow_duplicates=True) # Insert binary match column

    return Main_File, NewColName # Return modified dataframe and name

#  Generate the MEGA Report

def processMEGA( MetaACDataFilePaths, ValidationData, paymentSlugs,  filePath, BatchDate , WSDate, ACBatchStartTime, ACBatchEndTime):
    FileList = []
    ACData = generatePaymentReport(filePath, paymentSlugs )
    Funnels = ACData["PaymentFunnel"].unique()

    FunnelGrouping = {"Python": [ "10xTechies", "Python", "techies checkout"], "Excel" : ["Excel", "om checkout"], "AI" : ["AI", "ai checkout"]}

    if len(ACData) > 0:
        output_filename = f"ACData_{WSDate}.csv"
        FileList.append(output_filename)
        ACData.to_csv(output_filename, index=False, sep=",")

    #------------Meta AC Data Calculations---------------------

    MetaACData = pd.DataFrame()

    for file in MetaACDataFilePaths:
        data = pd.read_csv(file, sep=",")
        MetaACData = pd.concat([MetaACData, data], axis="rows")

    # Timezone Correction
    MetaACData["CreatedAt"] = MetaACData["CreatedAt"].astype('M8[s]')+pd.Timedelta(minutes=330)#.dt.strftime("%Y-%m-%d %H:%M:%S")

    MetaACData = MetaACData.sort_values(by = ["CreatedAt"], ascending=True )
    MetaACData.rename(columns = {"Payment Funnel": "PaymentFunnel"}, inplace =True)

    mergedACData = pd.concat([ACData, MetaACData], axis="rows", ignore_index=True)

    FileList = []
    #ExcludedData = []
    ACLeads = mergedACData
    st.write(f"ValidationData_{len(ValidationData)}")

    #pd.read_excel(ACOutput_FileName, sheet_name=Funnel)

    startDate, endDate, excludedStartDates, excludedEndDates = getDates("AI", BatchDate , WSDate, ACBatchStartTime, ACBatchEndTime)

    if startDate is not None:
        ACLeads = ACLeads[ (ACLeads["CreatedAt"].between(startDate, endDate) )]


    MFCombinations = ['Email',  'Phone Number']
    NonExoticSheetCombinations = ['Email',  'Phone Number']

    SumColNames = []

    ValidationData["Phone Number"] = ValidationData["Phone Number"].astype(str)
    ACLeads["Phone Number"] = ACLeads["Phone Number"].astype(str)

    CurrentFileSumColumns = [] #Stores all the columns for the summation
    for MFCol, CFCol  in zip(MFCombinations, NonExoticSheetCombinations):
        ACLeads, NewColName = CountIf(ACLeads, ValidationData, MFCol, CFCol, "Test")
        SumColNames.append(NewColName)
        CurrentFileSumColumns.append(NewColName)

    CurrentDateColName = "Total"

    ACLeads[CurrentDateColName] = ACLeads[CurrentFileSumColumns].sum(axis=1).gt(0).map({True: 'Matched', False: 'Unmatched'})
    MatchedLeads = ACLeads[ACLeads[CurrentDateColName] == "Matched"]
    ACLeads = ACLeads[ACLeads[CurrentDateColName] == "Unmatched"]

    #Drops all the column which was required for the countif function
    ACLeads.drop(columns=CurrentFileSumColumns+[CurrentDateColName ], inplace=True)

    st.write(f"Count of ACLeads before removing duplicates= {len(ACLeads)}")

    priority_order = ACLeads["Status"].unique().tolist()

    if "captured" in priority_order:
        priority_order.remove("captured")
        priority_order = ["captured"]+priority_order

    ACLeads["Status"] = pd.Categorical(ACLeads["Status"], categories=priority_order, ordered=True)

    ACLeads = ACLeads.sort_values(by=["Status", "Amount"], ascending=[True, False] )
    ACLeads["EmailLC"] = ACLeads["Email"].str.lower().str.strip()
    ACLeads = ACLeads.drop_duplicates(subset=["EmailLC"], keep="first" )
    ACLeads["Phone Number"] = ACLeads["Phone Number"].astype(float, errors = "ignore")
    ACLeads = ACLeads.drop_duplicates(subset=["Phone Number"], keep="first" )
    ACLeads.drop(columns=["EmailLC"], inplace=True)

    # Convert back to string before filtering to avoid Categorical comparison issues
    ACLeads["Status"] = ACLeads["Status"].astype(str)
    ACLeads = ACLeads[ACLeads.Status != "captured"]

    st.write(f"Count of ACLeads after removing duplicates= {len(ACLeads)}")

    ACLeads["Abandon Cart"] = "Yes"

    ACLeads = ACLeads.sort_values(by=["Status","CreatedAt"], ascending=[True, True])

    if "techies checkout" in ACLeads["Payment Slug"].unique():
        ACLeads = ACLeads[(~ACLeads.Amount.isin([114.46, 116.86]) )]

    ACLeads.to_excel(f"ACLeads_{WSDate}_Concat.xlsx", index=False)

    ACOutput_FileName = f"AC_Data_{WSDate}.xlsx"

    FunnelCount = pd.DataFrame(columns=["Funnel", "Count"])
    if len(ACLeads) > 0:
        with pd.ExcelWriter(ACOutput_FileName) as f:
            for fg in FunnelGrouping:
                sheet_name = f"{fg}_AC"
                OutputAC_DF = ACLeads.loc[ACLeads["PaymentFunnel"].isin(FunnelGrouping[fg]), :]
                if sheet_name == "Python_AC":
                    OutputAC_DF = OutputAC_DF[OutputAC_DF['Age Group'] != "Under 21"]
                FunnelCount.loc[len(FunnelCount), :]= [fg, len(OutputAC_DF)]
                OutputAC_DF.to_excel(f, sheet_name= sheet_name, index=False )
                output_filename = f"{sheet_name}.csv"
                OutputAC_DF.to_csv(output_filename, index=False, sep=",")
                FileList.append(output_filename)

    # if len(MatchedLeads) > 0:
    #     output_filename = f"ExcludedData_.csv"
    #     ExcludedData.append(output_filename)
    #     MatchedLeads = MatchedLeads.sort_values(by=["CreatedAt"], ascending=True)
    #     MatchedLeads.to_csv(output_filename, index=False, sep=",")

    return ACOutput_FileName, FunnelCount

def updateMegaSheet(sheet_id, file, sheet_name,  credential_Upload):
  # Authentication
  scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
  creds = ServiceAccountCredentials.from_json_keyfile_name(credential_Upload, scope)
  client = gspread.authorize(creds)

  workbook = client.open_by_key(sheet_id)

  try:  #Gets the batchname by removing splitting from the "W" part.
        df = pd.read_excel(file, sheet_name = sheet_name)

        MainFileBatches =  sheet_name

  except:
      MainFileBatches = None

  existing_sheet_titles = [ws.title for ws in workbook.worksheets()]
  #st.write(existing_sheet_titles)

  if MainFileBatches is not None:
    if MainFileBatches not in existing_sheet_titles:
        AddNewWS = workbook.add_worksheet(title=MainFileBatches, rows='100', cols='20')
        batchData = AddNewWS # Use the newly created worksheet object
    else:
        batchData = workbook.worksheet(MainFileBatches)
        batchData.clear()

  # Ensure set_with_dataframe receives the worksheet object
    set_with_dataframe(batchData, df)
    
    st.write(f"Process Completed for {sheet_name}.")
    return True

  else:
    st.write(file)


def check_session_state(sheet_id  ,sessionVarName , sheet_name , credential_Upload , clear):
        if sessionVarName not in st.session_state or clear:
            st.write(f"Downloading {sessionVarName}.. ")
            st.session_state[sessionVarName] = getSheet( sheet_id, sheet_name, credential_Upload)
            return st.session_state[sessionVarName]
        else:
            return  st.session_state[sessionVarName]
    

st.set_page_config("📊 MEGA Sheet - AC", layout="wide")
st.header("📊 MEGA Sheet - AC", divider=True, text_alignment="center")
WSDate =  str(st.date_input("Select the Next Sunday date",value=next_sunday()))
credential_Upload = st.file_uploader("Upload Credentials File", type = ["json"]) 
GdriveCredentials =  st.file_uploader("Upload GDrive File", type = ["json"]) 

MetaACData = st.file_uploader("Upload the files from Metabase", accept_multiple_files=True, type = ["csv"])

clearPreviousData = st.checkbox("Clear Data?")

if WSDate and MetaACData and GdriveCredentials and credential_Upload:
    genbtn = st.button("Generate Data", type="primary", on_click=None, use_container_width=True )

# Example usage:
    if genbtn:
        if "TotalFiles" in st.session_state:
            st.session_state.pop("TotalFiles") 

        if "upload_done" in st.session_state:
            st.session_state.pop("upload_done") 


        MetaACDataFilePaths = [save_upload(file) for file in MetaACData]

        CurrentDate = datetime.strptime(WSDate , "%Y-%m-%d")
        PreviousDate = CurrentDate - timedelta(days=8)
        ACBatchStartTime =  datetime.strptime(f"{str(PreviousDate.date())} 17:00:00", "%Y-%m-%d %H:%M:%S")
        ACBatchEndTime = ACBatchStartTime + timedelta(days = 7)

        credential_Upload = save_upload(credential_Upload)
        st.session_state["credential_Upload"] = credential_Upload

        getSheets = {"AI Exotic": ["AI", "AI BootcampPaid"], "SMAI Exotic": ["SMAI"], "PU Exotic": ["PU"], "AI Bootcamp":[ "AI","AI BootcampPaid" ]} 

        GdriveCredentials = save_upload(GdriveCredentials)

        with st.status("Processing..", expanded=True) as status:
            service = getGdriveService(GdriveCredentials)  # or getGdriveService(delegated_user="owner@yourdomain.com")
            filePath, service = getFilefromGdrive('0AHGO663tIOm5Uk9PVA', service, WSDate, clearPreviousData)
            
        # @title Downloading All Sheets
        # Remove existing file if it exists to avoid conflicts

            paymentSlugs = check_session_state("1v0UI5B4rkWJm3N8cbqnRCa4olvwV6-h-YC2mafNYnjU","paymentSlugs", WSDate, credential_Upload, clearPreviousData) 
 
            BatchDate = check_session_state("1szfXpbxy1lITxMU53e0TqlV_PjRVGv3OKTpI1wjoegk", "BatchDate", "BatchDate", credential_Upload, clearPreviousData)
 
            ExcludedTimings = check_session_state("1szfXpbxy1lITxMU53e0TqlV_PjRVGv3OKTpI1wjoegk", "ExcludedTimings", "ExcludedTimings", credential_Upload, clearPreviousData)
 
            MegaSheetInfo = check_session_state("1szfXpbxy1lITxMU53e0TqlV_PjRVGv3OKTpI1wjoegk", "MegaSheetInfo", "MegaSheetInfo", credential_Upload, clearPreviousData)
 
            condition = ((MegaSheetInfo["Date"] == WSDate) & (MegaSheetInfo["InUse"] == "Yes") )
            sheet_id = MegaSheetInfo.loc[condition,  "sheet_id" ].unique()[0]

            st.session_state["sheet_id"] = sheet_id

            AttendeeFile = r"MegaSheet.xlsx"
            url = f"https://drive.google.com/uc?id={sheet_id}"
            gdown.download(url, AttendeeFile, quiet=True)

            AdditionalSheetGrouping =  ["ExcludedDataAI Exotic_", "ExcludedDataPython_", "ExcludedDataAI_"]
            SheetGrouping = ['10xTechies', 'AI TV', 'AI', 'DRF', 'Excel', 'PU', 'Python', 'SMAI', 'AI Exotic', 'SMAI Exotic']+AdditionalSheetGrouping

            CurrentDate = datetime.strptime(WSDate , "%Y-%m-%d")
            PreviousDate = CurrentDate - timedelta(days=8)
            ACBatchStartTime =  datetime.strptime(f"{str(PreviousDate.date())} 17:00:00", "%Y-%m-%d %H:%M:%S")
            ACBatchEndTime = ACBatchStartTime + timedelta(days = 7)

            st.write(f"WSDate = {WSDate}, StartDate = {ACBatchStartTime}, EndDate = {ACBatchEndTime}")
            with zipfile.ZipFile(AttendeeFile) as z:
                with z.open('xl/workbook.xml') as f:
                    xml = f.read().decode('utf-8')

            ValidationDataSheetName = re.findall(r'<sheet[^>]*name="([^"]*)"', xml)

            ValidationDataSheetName = [x for x in SheetGrouping  if x in ValidationDataSheetName]

            if "ValidationData" not in st.session_state or clearPreviousData:
                ValidationData = pd.concat(
                    pd.read_excel(AttendeeFile, sheet_name=ValidationDataSheetName),  # returns dict
                    ignore_index=True)

                st.session_state["ValidationData"] = ValidationData
            else:
                ValidationData = st.session_state["ValidationData"]

            status.update(label="Completed!",expanded=False)

            ACOutput_FileName, FunnelCount = processMEGA(MetaACDataFilePaths,ValidationData, paymentSlugs,  filePath,  BatchDate , 
                                                         WSDate, ACBatchStartTime, ACBatchEndTime)

        st.dataframe(MegaSheetInfo.loc[condition, ["Date", "sheet_id"]] , hide_index=True)

        st.dataframe(FunnelCount, hide_index=True)
        

        st.session_state["TotalFiles"] = ACOutput_FileName

        ACOutput_FileName = rf"AC_Data_{WSDate}.xlsx"
             

        with open(ACOutput_FileName, "rb") as f:
            st.download_button(
                label="Save MEGA file",
                data=f.read(),
                file_name=ACOutput_FileName,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click="ignore"
            )
    #Unmatched_Slugs.to_excel(f, sheet_name="Unmatched_Slugs", index=False)
    
        st.info(sheet_id)

    if "TotalFiles" in st.session_state:
        if "upload_done" not in st.session_state:
            st.session_state["upload_done"] = False

        ACOutput_FileName = st.session_state["TotalFiles"]
        sheet_id = st.session_state["sheet_id"]
        credential_Upload = st.session_state["credential_Upload"]
        
        if not st.session_state["upload_done"]:
            upload = st.button("Upload Data?", type="primary", key="upload")
            if upload:
                with st.status("Uploading..", expanded=True) as status:
                    with pd.ExcelFile(ACOutput_FileName) as f:
                        sheet_names = f.sheet_names
                        for sheet in stqdm(sheet_names):
                            updateMegaSheet(sheet_id, ACOutput_FileName, sheet, credential_Upload) 

                        status.update(label="Upload Complete!!",expanded=False)

                        st.success("Upload Completed!")

                        st.session_state["upload_done"] = True  # mark upload as complete
                        #os.unlink(ACOutput_FileName)
                        st.rerun()
        else:
                st.success("Upload Completed!")

else:
    with st.status("Paste this code in Metabase", expanded=True):
        FunnelsList = ["AI for Techies Checkout", "BE10X Checkout", "Office Master Checkout"]
        FunnelCheckBox = st.selectbox("Select a Funnel to get code", FunnelsList)
        st.code(generate_code(FunnelCheckBox, WSDate))


