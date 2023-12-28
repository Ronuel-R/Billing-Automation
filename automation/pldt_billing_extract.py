from time import sleep
from datetime import datetime

from extracting_helper import encode

import os, requests, logging, pyperclip, shutil, csv, pyautogui, pygetwindow as gw, webbrowser,socket

def get_value_for_type(file_path_master_file, target_account):
    with open(file_path_master_file, 'r') as csvfile:

        csv_reader = csv.reader(csvfile)
        
        header = next(csv_reader, None)
        
        for row in csv_reader:

            if row and row[0].strip().lower() == target_account.lower():
                name_value = row[1].strip()
                return name_value

    logging.info("Account Number with {} not found in the Master File.".format(target_account))
    return 'N/A'

def get_value_for_cost_center(file_path_master_file, target_account):
    with open(file_path_master_file, 'r') as csvfile:

        csv_reader = csv.reader(csvfile)
        
        header = next(csv_reader, None)
        
        for row in csv_reader:

            if row and row[0].strip().lower() == target_account.lower():

                cost_center_value = row[2].strip()
                return cost_center_value

    logging.info("Account Number with {} not found in the Master File.".format(target_account))
    return 'N/A'
shared_folder_path = r'\\172.18.118.9\d\ADMINISTRATION\PLDT'
pre_requisite_path = r'\\172.18.118.9\d\ADMINISTRATION\PLDT\PLDT_PREREQUISITE'
directory_path = r'\\172.18.118.9\d\ADMINISTRATION\PLDT\PLDT_BILLING'
logging.basicConfig(filename=shared_folder_path +'/LOGS/billings.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',filemode='a')
logging.info("Program started.")

current_year = str(datetime.now().year)
fields = ['SOA Number','Account','Billing' ,'Amount To Pay','Due Date']  
rows_list = []
processed = 0
errors = 0
date_initial = []
rows_list_extracted = []
rows_extracted_dict = {}
rows_list_error = []
rows_error_dict = {}
try:
    all_files = os.listdir(directory_path)
    pdf_files = [file for file in all_files if file.lower().endswith('.pdf')]
    for pdf_file in pdf_files:
        webbrowser.open_new('file://' + directory_path + '\\'+ pdf_file)
        sleep(5)
        window_title = pdf_file
        window = gw.getWindowsWithTitle(window_title)[0]
        if window:
            window.activate()
        sleep(3)
        try:
            res_amount = pyautogui.locateCenterOnScreen(pre_requisite_path + "\current_balance.png", grayscale=True, confidence=.8) #type: ignore
            if res_amount:
                res_amount_value = res_amount._replace(x=res_amount.x + 10,y=res_amount.y + 15)
                pyautogui.moveTo(res_amount_value)
                pyautogui.doubleClick()
                pyautogui.hotkey("ctrl","c")
                copied_amount_value = pyperclip.paste()
                sleep(1)
  
            res_account_num = pyautogui.locateCenterOnScreen(pre_requisite_path + "/account_number.png", grayscale=True, confidence=.8) #type: ignore
            if res_account_num:
                res_account_num_value = res_account_num._replace(y=res_account_num.y + 15)
                pyautogui.moveTo(res_account_num_value)
                pyautogui.doubleClick()
                pyautogui.hotkey("ctrl","c")
                copied_account_value = pyperclip.paste()
                sleep(1)

            res_billing = pyautogui.locateCenterOnScreen(pre_requisite_path + "/bill_period.png", grayscale=True, confidence=.8) #type: ignore
            if res_billing:
                res_billing_value = res_billing._replace(x=res_billing.x - 35,y=res_billing.y + 15)
                pyautogui.moveTo(res_billing_value)
                pyautogui.dragTo(res_billing_value.x + 65,res_billing_value.y,1,button='left')
                pyautogui.mouseUp()
                sleep(1)
                pyautogui.hotkey("ctrl","c")
                copied_billing_value = pyperclip.paste()
                sleep(1)
                
            res_soa = pyautogui.locateCenterOnScreen(pre_requisite_path + "/soa_number.png", confidence=.7) #type: ignore
            if res_soa:
                res_soa_value = res_soa._replace(y=res_soa.y + 15)
                pyautogui.moveTo(res_soa_value)
                pyautogui.doubleClick()
                pyautogui.hotkey("ctrl","c")
                copied_soa_value = pyperclip.paste()
                cleaned_soa_value = copied_soa_value.lstrip('0')
                sleep(1)
                
            res_due_date = pyautogui.locateCenterOnScreen(pre_requisite_path + "/due_date.png", confidence=.8) #type: ignore
            if res_due_date:
                res_due_date_value = res_due_date._replace(x=res_due_date.x - 40,y=res_due_date.y + 15)
                pyautogui.moveTo(res_due_date_value)
                pyautogui.dragTo(res_due_date_value.x + 75,res_due_date_value.y,1,button='left')
                pyautogui.mouseUp()
                sleep(1)
                pyautogui.hotkey("ctrl","c")
                raw_due_value = pyperclip.paste()
                cleaned_due_value = datetime.strptime(raw_due_value, "%b %d, %Y")
                copied_due_value = cleaned_due_value.strftime("%m/%d/%Y")
                if copied_soa_value == copied_due_value: #type:ignore
                    copied_due_value = 'N/A'
                sleep(1)
                
            list = [cleaned_soa_value,copied_account_value,copied_billing_value,copied_amount_value,copied_due_value]#type: ignore
            found = any(copied_account_value in inner_list for inner_list in rows_list)#type:ignore
            if not found: #type:ignore
                if rows_list == []:
                    rows_list.append(list)
                    logging.info("{} | Success.".format(pdf_file))
                    date_initial.append(copied_billing_value) #type:ignore
                    pyautogui.hotkey("ctrl","w")
                    source_file = directory_path + '/' + pdf_file
                    destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + copied_billing_value + '/' #type:ignore
                    if not os.path.exists(destination_folder):
                        os.makedirs(destination_folder)
                    shutil.move(source_file, destination_folder)
                    logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

                elif copied_billing_value not in date_initial: #type:ignore
                    rows_list.append(list)
                    logging.info("{} | Success.".format(pdf_file))
                    pyautogui.hotkey("ctrl","w")
                    date_initial.append(copied_billing_value) #type:ignore
                    source_file = directory_path + '/' + pdf_file
                    destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + copied_billing_value + '/' #type:ignore
                    if not os.path.exists(destination_folder):
                        os.makedirs(destination_folder)
                    shutil.move(source_file, destination_folder)
                    logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

                else:
                    rows_list.append(list)
                    logging.info("{} | Success.".format(pdf_file))
                    pyautogui.hotkey("ctrl","w")
                    source_file = directory_path + '/' + pdf_file
                    destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + copied_billing_value + '/' #type:ignore
                    if not os.path.exists(destination_folder):
                        os.makedirs(destination_folder)
                    shutil.move(source_file, destination_folder)
                    logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))
            else:
                logging.info("{} | Already Existing.".format(pdf_file))
                pyautogui.hotkey("ctrl","w")
                source_file = directory_path + '/' + pdf_file
                destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + copied_billing_value + '/' #type:ignore
                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                shutil.move(source_file, destination_folder)
                logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))
                
        except Exception as e:
            logging.error("{} | Error: {}".format(pdf_file, e))
            errors += 1
            pyautogui.hotkey("ctrl","w")
            source_file = directory_path + '/' + pdf_file
            destination_folder = shared_folder_path + '/LOGS/ERROR/'
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            shutil.move(source_file, destination_folder)
            logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

    original_data = rows_list.copy()
    for date_i in date_initial:
        for sublist in original_data:
            if "N/A" in sublist[2]:
                rows_list_error.append(sublist)
                rows_list.remove(sublist)
                
            else:
                rows_list_extracted.append(sublist)
                rows_list.remove(sublist)
                    
        rows_extracted_dict[date_i] = rows_list_extracted
        rows_error_dict[date_i] = rows_list_error
        rows_list_extracted = []
        rows_list_error = []

    
    for i, date_i in enumerate(date_initial):
        if len(rows_extracted_dict[date_i]) != 0:
            filename = str(date_i) + ".csv"
            destination_folder = shared_folder_path + '/EXTRACTED/' 

            file_path = os.path.join(destination_folder, filename)

            with open(file_path, 'w') as csvfile:  
                csvwriter = csv.writer(csvfile)  
                csvwriter.writerow(fields)
                csvwriter.writerows(rows_extracted_dict[date_i])
                logging.info("CSV File created.")

    
    for i, date_i in enumerate(date_initial):
        if len(rows_error_dict[date_i]) != 0:
            filename = str(date_i) + ".csv"
            destination_folder = shared_folder_path + '/LOGS/ERROR/' 

            file_path = os.path.join(destination_folder, filename)

            with open(file_path, 'w') as csvfile:  
                csvwriter = csv.writer(csvfile)  
                csvwriter.writerow(fields)
                csvwriter.writerows(rows_error_dict[date_i])
                logging.info("Error CSV File created.")
        
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)

    url = 'insert_api_here' if ip_address != '172.18.118.5' else 'insert_api_here'
    token = encode('automation', 'pldt') # type: ignore
    headers = {
    'Authorization': token,
    }

    response = requests.request('GET',url, headers=headers)
    logging.info("Response Code: {}".format(response.status_code))
    logging.info("Program finished.\n{}".format('=' * 145))

except Exception as e:
    if e:
        logging.error(e)
        logging.info("Program finished.\n{}".format('=' * 145))
    else:
        logging.info("No files to be processed.")
        logging.info("Program finished.\n{}".format('=' * 145))
