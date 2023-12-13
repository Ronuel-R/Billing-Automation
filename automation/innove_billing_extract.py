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
shared_folder_path = r'\\172.18.118.9\d\ADMINISTRATION\INNOVE'
pre_requisite_path = r'\\172.18.118.9\d\ADMINISTRATION\INNOVE\INNOVE_PREREQUISITE'
directory_path = r'\\172.18.118.9\d\ADMINISTRATION\INNOVE\INNOVE_BILLING'
logging.basicConfig(filename=shared_folder_path +'/LOGS/billings.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',filemode='a')
logging.info("Program started.")

current_year = str(datetime.now().year)
# directory_path = r'C:\Users\Ronuel\Documents\Innove_Billing_Automation\Innove_Billing'
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
            try:
                res_amount_first = pyautogui.locateCenterOnScreen(pre_requisite_path + "\current_balance_first.png", grayscale=True, confidence=.8) #type: ignore
                if res_amount_first:
                    res_amount_value = res_amount_first._replace(x=res_amount_first.x + 150)
                    pyautogui.moveTo(res_amount_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_amount_value = pyperclip.paste()
                    sleep(1)
            except:
                logging.warning('PDF file does not contain "current_balance_first.png".')
                res_amount_second = pyautogui.locateCenterOnScreen(pre_requisite_path + "\current_balance_second.png", grayscale=True, confidence=.8) #type: ignore
                if res_amount_second:
                    res_amount_value = res_amount_second._replace(x=res_amount_second.x + 270)
                    pyautogui.moveTo(res_amount_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_amount_value = pyperclip.paste()
                    sleep(1) 
                else:
                    raise Exception("Amount image not found on the screen")
            try:    
                res_account_num_first = pyautogui.locateCenterOnScreen(pre_requisite_path + "/account_number_first.png", grayscale=True, confidence=.8) #type: ignore
                if res_account_num_first:
                    res_account_num_first_value = res_account_num_first._replace(x=res_account_num_first.x + 150)
                    pyautogui.moveTo(res_account_num_first_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_account_value = pyperclip.paste()
                    sleep(1)
            except:
                logging.warning('PDF file does not contain "account_number_first.png".')
                res_account_num_second = pyautogui.locateCenterOnScreen(pre_requisite_path + "/account_number_second.png", grayscale=True, confidence=.8) #type: ignore
                if res_account_num_second:
                    res_account_num_value = res_account_num_second._replace(y=res_account_num_second.y + 15)
                    pyautogui.moveTo(res_account_num_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_account_value = pyperclip.paste()
                    sleep(1) 
                else:
                    raise Exception("Amount image not found on the screen")
            try:
                res_billing_first = pyautogui.locateCenterOnScreen(pre_requisite_path + "/bill_period_first.png", grayscale=True, confidence=.8) #type: ignore
                if res_billing_first:
                    res_billing_first_value = res_billing_first._replace(x=res_billing_first.x + 80)
                    pyautogui.moveTo(res_billing_first_value)
                    pyautogui.dragTo(res_billing_first_value.x + 130,res_billing_first_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_billing_value = pyperclip.paste()
                    cleaned_string = copied_billing_value.replace('\r', '')
                    sleep(1)
            except:
                logging.warning('PDF file does not contain "bill_period_first.png".')
                res_billing_second = pyautogui.locateCenterOnScreen(pre_requisite_path + "/bill_period_second.png", grayscale=True, confidence=.8) #type: ignore
                if not res_billing_second:
                    raise Exception("Amount image not found on the screen")
                if res_billing_second:
                    res_billing_second_value = res_billing_second._replace(y=res_billing_second.y + 15 ,x=res_billing_second.x + 55)
                    pyautogui.moveTo(res_billing_second_value)
                    pyautogui.dragTo(res_billing_second_value.x - 140,res_billing_second_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_billing_value = pyperclip.paste()
                    start_date_str, end_date_str = copied_billing_value.split(" to ")
                    start_date_obj = datetime.strptime(start_date_str, "%m/%d/%y")
                    end_date_obj = datetime.strptime(end_date_str, "%m/%d/%y")
                    formatted_start_date = start_date_obj.strftime("%m-%d-%y")
                    formatted_end_date = end_date_obj.strftime("%m-%d-%y")
                    cleaned_string = "{} to {}".format(formatted_start_date,formatted_end_date)
                    sleep(1) 
                else:
                    raise Exception("Amount image not found on the screen")
            try:
                res_soa_first = pyautogui.locateCenterOnScreen(pre_requisite_path + "/soa_number_first.png", confidence=.7) #type: ignore
                if res_soa_first:
                    res_soa_first_value = res_soa_first._replace(x=res_soa_first.x + 40)
                    pyautogui.moveTo(res_soa_first_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_soa_value = pyperclip.paste()
                    editted_soa_value = copied_soa_value.lstrip('I')
                    cleaned_soa_value = editted_soa_value.lstrip('0')
                    sleep(1)
            except:
                logging.warning('PDF file does not contain "soa_number_first.png".')
                res_soa_second = pyautogui.locateCenterOnScreen(pre_requisite_path + "/soa_number_second.png", confidence=.7) #type: ignore
                if res_soa_second:
                    res_soa_second_value = res_soa_second._replace(x=res_soa_second.x + 30)
                    pyautogui.moveTo(res_soa_second_value)
                    pyautogui.doubleClick()
                    pyautogui.hotkey("ctrl","c")
                    copied_soa_value = pyperclip.paste()
                    editted_soa_value = copied_soa_value.lstrip('IN')
                    cleaned_soa_value = editted_soa_value.lstrip('0')
                    sleep(1)
                else:
                    raise Exception("Amount image not found on the screen")
            try:   
                res_due_date_first = pyautogui.locateCenterOnScreen(pre_requisite_path + "/due_date_first.png", confidence=.8) #type: ignore
                if res_due_date_first:
                    res_due_date_first_value = res_due_date_first._replace(x=res_due_date_first.x + 25)
                    pyautogui.moveTo(res_due_date_first_value)
                    pyautogui.dragTo(res_due_date_first_value.x + 60,res_due_date_first_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    raw_due_value = pyperclip.paste()
                    cleaned_due_value = datetime.strptime(raw_due_value, "%b %d, %Y")
                    copied_due_value = cleaned_due_value.strftime("%m/%d/%Y")
                    if copied_soa_value == copied_due_value: #type:ignore
                        copied_due_value = 'N/A'
                    sleep(1)
            except:
                logging.warning('PDF file does not contain "due_date_first.png".')
                res_due_date_second = pyautogui.locateCenterOnScreen(pre_requisite_path + "\due_date_second.png", confidence=.8) #type: ignore
                if res_due_date_second:
                    res_due_date_second_value = res_due_date_second._replace(y=res_due_date_second.y + 15,x=res_due_date_second.x - 15)
                    pyautogui.moveTo(res_due_date_second_value)
                    pyautogui.dragTo(res_due_date_second_value.x + 50,res_due_date_second_value.y,1,button='left')
                    pyautogui.mouseUp()
                    sleep(1)
                    pyautogui.hotkey("ctrl","c")
                    copied_due_value = pyperclip.paste()
                    if copied_soa_value == copied_due_value: #type:ignore
                        copied_due_value = 'N/A'
                    sleep(1)
                else:
                    raise Exception("Amount image not found on the screen")
                
            list = [cleaned_soa_value,copied_account_value,cleaned_string,copied_amount_value,copied_due_value]#type: ignore

            if rows_list == []:
                rows_list.append(list)
                logging.info("{} | Success.".format(pdf_file))
                date_initial.append(cleaned_string) #type:ignore
                pyautogui.hotkey("ctrl","w")
                source_file = directory_path + '/' + pdf_file
                destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + cleaned_string + '/' #type:ignore
                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                shutil.move(source_file, destination_folder)
                logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

            elif cleaned_string not in date_initial: #type:ignore
                rows_list.append(list)
                logging.info("{} | Success.".format(pdf_file))
                pyautogui.hotkey("ctrl","w")
                date_initial.append(cleaned_string) #type:ignore
                source_file = directory_path + '/' + pdf_file
                destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + cleaned_string + '/' #type:ignore
                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                shutil.move(source_file, destination_folder)
                logging.info("{} | Moved to | {}.".format(pdf_file,destination_folder))

            else:
                rows_list.append(list)
                logging.info("{} | Success.".format(pdf_file))
                pyautogui.hotkey("ctrl","w")
                source_file = directory_path + '/' + pdf_file
                destination_folder = shared_folder_path + '/ARCHIVED/' + current_year + '/' + cleaned_string + '/' #type:ignore
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
                logging.info("CSV File created.")
        
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)

    url = 'http://127.0.0.1:8000/rfp/automate/innove/' if ip_address != '172.18.118.5' else 'http://portal.puremart.ph/rfp/automate/innove/'
    token = encode('innove', 'filename') # type: ignore
    headers = {
    'Authorization': 'Bearer {}'.format(token),
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