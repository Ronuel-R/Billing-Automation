import base64, csv, inflect
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from datetime import datetime, timedelta
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from rfp.models import InnoveMasterList, PostpaidMasterList

def make_row_postpaid_master_list(data):

        arr = (data['id'],
               data['account_number'],
               data['name'],
               data['cost_center'],
               )
               
        return arr

def make_row_innove_master_list(data):

        arr = (data['id'],
               data['account_number'],
               data['type'],
               data['cost_center'],
               )
               
        return arr

def paginate_page(length,start_page,serializer,records_total):
    lists_data = []
    if int(length) < 0:
        paginator = Paginator(serializer.data, records_total) # type: ignore
        page_number = 1
    else:
        paginator = Paginator(serializer.data, int(length)) # type: ignore
        page_number = int(start_page) / int(length) + 1
    
    try:
        lists_data = paginator.page(page_number).object_list # type: ignore
    except PageNotAnInteger:
        lists_data = paginator.page(1).object_list
    except EmptyPage:
        lists_data = paginator.page(1).object_list
    except:
        pass
    
    return lists_data

def get_reference_values(csvfile_path,target):
    reference = []
    with open(csvfile_path, 'r') as csvfile:
        
        csv_reader = csv.reader(csvfile)
        
        header = next(csv_reader, None)

        for index, row in enumerate(csv_reader, start=1):
            if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                value = row[0].strip()
                reference.append(value)

    return reference

def get_particulars_values(csvfile_path,target):
    particulars = []
    with open(csvfile_path, 'r') as csvfile:
        csv_reader = csv.reader(csvfile)

        header = next(csv_reader, None)

        for index, row in enumerate(csv_reader, start=1):
            if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                if target == 'postpaid':
                    if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                        account_value = row[1].strip()
                        billing_start_value = row[2].strip()
                        billing_end_value = row[5].strip()
                        if target == 'postpaid':
                            postpaid_obj = PostpaidMasterList.objects.get(account_number=account_value)
                            name_value = postpaid_obj.name
                        particulars.append("Account Number: {} - {}\nBilling Period: {} to {}".format(account_value, name_value, billing_start_value, billing_end_value)) # type: ignore
                elif target == 'innove':
                    account_value = row[1].strip()
                    billing_value = row[2].strip()
                    if target == 'innove':
                        try:
                            innove_obj = InnoveMasterList.objects.get(account_number=account_value)
                            name_value = innove_obj.type
                        except:
                            name_value = "N/A"
                    
                    particulars.append("Account Number: {} - {}\nBilling Period: {}".format(account_value, name_value, billing_value)) # type: ignore
    return particulars

def get_cost_center_values(csvfile_path,target):
    cost_center = []
    with open(csvfile_path, 'r') as csvfile:
        csv_reader = csv.reader(csvfile)

        header = next(csv_reader, None)

        for index, row in enumerate(csv_reader, start=1):
            if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                account_value = row[1].strip()

                if target == 'postpaid':
                    postpaid_obj = PostpaidMasterList.objects.get(account_number=account_value)
                    cost_center_value = postpaid_obj.cost_center

                elif target == 'innove':
                    try:
                        innove_obj = InnoveMasterList.objects.get(account_number=account_value)
                        cost_center_value = innove_obj.cost_center
                    except:
                        cost_center_value = "No location set for account number {}".format(account_value)

                cost_center.append(cost_center_value) # type: ignore

    return cost_center

def get_amount_values(csvfile_path,target):
    amount = []
    with open(csvfile_path, 'r') as csvfile:
        csv_reader = csv.reader(csvfile)

        header = next(csv_reader, None)

        for index, row in enumerate(csv_reader, start=1):
            if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                if target == 'postpaid':
                    value = row[3].strip()
                    amount_float = value.replace(',', '')
                    amount.append(amount_float)
                if target == 'innove':
                    if InnoveMasterList.objects.filter(account_number=row[1]).first():
                        value = row[3].strip()
                        amount_float = value.replace(',', '')
                        amount.append(amount_float)
                    else:
                        amount.append("0")
    return amount

def get_amt_word(amount,target):
    amt_words = []
    total_amt = 0.00
    p = inflect.engine()

    for amt in amount:
        amt_without_comma = amt.replace(',', '')
        total_amt += float(amt_without_comma)

    amt_words = p.number_to_words(int(total_amt), andword='')

    cents_part = '{:.2f}'.format(total_amt % 1).split('.')[1]


    output = "{} PESOS AND {}/100 ONLY".format(amt_words.upper(), cents_part)
    return output

def get_total_amt(amount,target):
    total_amt = 0

    for amt in amount:
        amt_without_comma = amt.replace(',', '')
        total_amt += round(float(amt_without_comma), 2)
        total_amt = round(total_amt, 2)
    return total_amt

def get_due_date(csvfile_path,target):
    formatted_due_date = ""
    current_date = datetime.now()
    try:
        with open(csvfile_path, 'r') as csvfile:
            csv_reader = csv.reader(csvfile)

            header = next(csv_reader, None)

            for index, row in enumerate(csv_reader, start=1):
                if index % 2 == 0 and index >= 2 and row and len(row) > 1:
                    due_date_str = row[4].strip()
                    if due_date_str != "N/A":
                        due_date = datetime.strptime(due_date_str, "%m/%d/%y")

                        if due_date == current_date:
                            due_date_in = due_date + timedelta(days=2)
                            formatted_due_date = due_date_in.strftime("%m-%d-%Y")
                            break
                        elif due_date < current_date: 
                            due_date = current_date + timedelta(days=2)
                            formatted_due_date = due_date.strftime("%m-%d-%Y")
                            break
                        elif due_date > current_date:
                            formatted_due_date = due_date.strftime("%m-%d-%Y")
                            break
                else:
                    raise ValueError("No Valid Due date matched")

    except Exception as e:
        due_date = current_date + timedelta(days=2)
        formatted_due_date = due_date.strftime("%m-%d-%Y")

    return str(formatted_due_date)