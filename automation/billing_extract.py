from time import sleep
from datetime import datetime

from extracting_helper import encode

import os, requests, logging, pyperclip, shutil, csv, pyautogui, pygetwindow as gw, webbrowser,socket

def get_value_for_name(file_path_master_file, target_account):
    with open(file_path_master_file, 'r') as csvfile:

        csv_reader = csv.reader(csvfile)
        
        header = next(csv_reader, None)
        
        for row in csv_reader:

            if row and row[1].strip().lower() == target_account.lower():
                name_value = row[2].strip()
                return name_value

    logging.warning("Account Number with {} not found in the Master File.".format(target_account))
    return 'N/A'

def get_value_for_cost_center(file_path_master_file, target_account):
    with open(file_path_master_file, 'r') as csvfile:

        csv_reader = csv.reader(csvfile)
        
        header = next(csv_reader, None)
        
        for row in csv_reader:

            if row and row[1].strip().lower() == target_account.lower():

                cost_center_value = row[3].strip()
                return cost_center_value

    logging.warning("Account Number with {} not found in the Master File.".format(target_account))
    return 'N/A'

shared_folder_path = r'\\172.18.118.9\d\ADMINISTRATION\GLOBE'
pre_requisite_path = r'\\172.18.118.9\d\ADMINISTRATION\GLOBE\GLOBE_PREREQUISITE'
directory_path = r'\\172.18.118.9\d\ADMINISTRATION\GLOBE\GLOBE_POSTPAID'
logging.basicConfig(filename=shared_folder_path +'/Logs/billings.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',filemode='a')
logging.info("Program started.")
current_year = str(datetime.now().year)

try:
    folders = [f for f in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, f))]
    for monthly_folder in folders:
        directory_path_to_pdf = os.path.join(directory_path, monthly_folder, 'BC27')
        logging.info("{} | Successfully located.".format(directory_path_to_pdf))
        all_files = os.listdir(directory_path_to_pdf)
        pdf_files = [file for file in all_files if file.lower().endswith('.pdf')]
        rows_list = []
        rows_list_error = []
        original_data = []
        fields = ['SOA Number','Account','Billing Start' ,'Amount To Pay','Due Date','Billing End']  
        date_initial = ""

        for pdf_file in pdf_files:
            webbrowser.open_new('file://' + directory_path_to_pdf + '\\'+ pdf_file)
            sleep(5)
            window_title = pdf_file
            window = gw.getWindowsWithTitle(window_title)[0]
            if window:
                window.activate()
            sleep(3)
            try:
                res_amount = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_Amount_To_Pay.png", grayscale=True, confidence=.8) # type: ignore
                if not res_amount:
                    raise Exception("Amount image not found on the screen")
                if res_amount:
                    res_amount_value = res_amount._replace(x=res_amount.x + 280)
                    pyautogui.moveTo(res_amount_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_amount_value = pyperclip.paste()
                    sleep(1) 
                res_account_num = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_Account_Number.png", grayscale=True, confidence=.8) # type: ignore
                if not res_account_num:
                    raise Exception("Amount image not found on the screen")
                if res_account_num:
                    res_account_num_value = res_account_num._replace(y=res_account_num.y + 15)
                    pyautogui.moveTo(res_account_num_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_account_value = pyperclip.paste()
                    sleep(1)

                res_billing_end = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_Billing_Period.png", grayscale=True, confidence=.8) # type: ignore
                if not res_billing_end:
                    raise Exception("Amount image not found on the screen")
                if res_billing_end:
                    res_billing_end_value = res_billing_end._replace(y=res_billing_end.y + 15 ,x=res_billing_end.x + 55)
                    pyautogui.moveTo(res_billing_end_value)
                    pyautogui.dragTo(res_billing_end_value.x + 55,res_billing_end_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_billing_end_value = pyperclip.paste()
                    sleep(1)
                
                res_billing_start = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_Billing_Period.png", grayscale=True, confidence=.8) # type: ignore
                if not res_billing_start:
                    raise Exception("Amount image not found on the screen")
                if res_billing_start:
                    res_billing_start_value = res_billing_start._replace(y=res_billing_start.y + 15 ,x=res_billing_start.x + 25)
                    pyautogui.moveTo(res_billing_start_value)
                    pyautogui.dragTo(res_billing_start_value.x - 70,res_billing_start_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_billing_start_value = pyperclip.paste()
                    sleep(1)
                
                res_soa = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_SOA.png", confidence=.7) # type: ignore
                if not res_soa:
                    raise Exception("Amount image not found on the screen")
                if res_soa:
                    res_soa_value = res_soa._replace(x=res_soa.x + 30)
                    pyautogui.moveTo(res_soa_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_soa_value = pyperclip.paste()
                    cleaned_soa_value = copied_soa_value.lstrip('0')
                    sleep(1)
                    
                res_due_date = pyautogui.locateCenterOnScreen(pre_requisite_path + "\Globe_Due_Date.png", confidence=.8) # type: ignore
                if res_due_date:
                    res_due_date_value = res_due_date._replace(y=res_due_date.y + 15,x=res_due_date.x - 15)
                    pyautogui.moveTo(res_due_date_value)
                    pyautogui.dragTo(res_due_date_value.x + 55,res_due_date_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_due_value = pyperclip.paste()
                    if copied_soa_value == copied_due_value: # type: ignore
                        copied_due_value = 'N/A'
                    sleep(1)

                list = [cleaned_soa_value,copied_account_value,copied_billing_start_value,copied_amount_value,copied_due_value,copied_billing_end_value] # type: ignore

                rows_list.append(list)
                logging.info("{} | Success.".format(pdf_file))
                date_initial = copied_billing_start_value # type: ignore
                pyautogui.hotkey("ctrl","w")
                date = datetime.strptime(date_initial, "%m/%d/%y")
                    
            except Exception as e:
                logging.error("{} | Error.".format(pdf_file))
                pyautogui.hotkey("ctrl","w")
                source_file = directory_path_to_pdf + '/' +  pdf_file
                destination_folder = shared_folder_path + '/LOGS/ERROR/' + monthly_folder
                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                shutil.move(source_file, destination_folder)
                logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

        original_data = rows_list.copy()
        for sublist in original_data:
            if "N/A" in sublist[2]:
                rows_list_error.append(sublist)
                rows_list.remove(sublist)

        if len(rows_list) != 0:
            filename = monthly_folder + ".csv"
            destination_folder = shared_folder_path + '/Extracted/' 

            file_path = os.path.join(destination_folder, filename)

            with open(file_path, 'w') as csvfile:  
                csvwriter = csv.writer(csvfile)  
                csvwriter.writerow(fields)
                csvwriter.writerows(rows_list)
                logging.info("CSV File created.")

        if len(rows_list_error) != 0:
            filename_error = monthly_folder + ".csv"
            destination_folder = shared_folder_path + '/LOGS/ERROR/'

            file_path = os.path.join(destination_folder, filename_error)

            with open(file_path, 'w') as csvfile:  
                csvwriter = csv.writer(csvfile)  
                csvwriter.writerow(fields)
                csvwriter.writerows(rows_list_error)
                logging.info("CSV File created.")
        
        folder_path_from = directory_path + '/{}'.format(monthly_folder)
        destination_folder = shared_folder_path + '/ARCHIVED/' + current_year 

        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        else:
            pass
        
        if os.path.exists(folder_path_from):
            shutil.move(folder_path_from, destination_folder)
        else:
            pass

    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)

    url = 'insert_api_here' if ip_address != '172.18.118.5' else 'insert_api_here'
    token = encode('automation', 'postpaid') # type: ignore
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
