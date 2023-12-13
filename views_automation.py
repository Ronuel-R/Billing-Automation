from django.http import HttpResponseRedirect,JsonResponse
import os, sys, shutil
from rfp.forms import *
from rfp.models import *
from django.db import transaction
from datetime import date, datetime
from rest_framework.response import Response
from tta.utils import *
from tds.models import *
from django.template.loader import render_to_string
from .automation_helper import *
from main.models import Location, Department
from django.contrib.auth.decorators import login_required
from main.mailer import mail_automation,automation_template_html
from rest_framework.permissions import IsAuthenticated 
from rest_framework.decorators import api_view, permission_classes
from rfp.serializers import *
import pandas as pd

@transaction.atomic
def rfp_automate_postpaid(request,*args,**kwargs):

    token = request.headers.get('Authorization')

    if not token:
        return JsonResponse({'error': 'Token not provided'}), 401
    
    assigned_user = Profile.objects.get(selectedForAutomation = True)
    info = {}
    data = dict()
    directory_path = r'\\172.18.118.9\d\ADMINISTRATION\GLOBE\EXTRACTED'
    destination_folder_path = r'\\172.18.118.9\d\ADMINISTRATION\GLOBE'
    RFPDPOST = ""
    current_year = str(datetime.now().year)
    all_files = os.listdir(directory_path)
    csv_files = [file for file in all_files if file.lower().endswith('.csv')]
    file_name_list = []
    total_amount_list =[]
    payee_list = []
    dates_list = []
    due_date_list = []
    status_list = []
    error_list = []
    for csv_file in csv_files:
        file_name = os.path.splitext(os.path.basename(csv_file))[0]
        
        csvfile_path = directory_path + '/{}'.format(csv_file)

        info['REFERENCE'] = get_reference_values(csvfile_path,'postpaid') #type:ignore
        info['PARTICULARS'] = get_particulars_values(csvfile_path,'postpaid') #type:ignore
        info['AMOUNT'] = get_amount_values(csvfile_path,'postpaid') #type:ignore
        info['COSTCENTER'] = get_cost_center_values(csvfile_path,'postpaid') #type:ignore
        
        info['DUEDATE'] = str(get_due_date(csvfile_path,'postpaid')) #type:ignore
        info['AMTINWORDS'] = get_amt_word(info['AMOUNT'],'postpaid') #type:ignore
        info['TOTALAMOUNT'] = get_total_amt(info['AMOUNT'],'postpaid') #type:ignore
        info['PAYTO'] = "GLOBE TELECOM, INC."
        file_name_list.append(str(file_name))
        total_amount_list.append(str(info['TOTALAMOUNT']))
        payee_list.append(str(info['PAYTO']))
        dates_list.append(str(date.today().strftime("%m-%d-%Y")))
        due_date_list.append(str(info['DUEDATE']))
        
        
        try:
            form = Create_RFPForm(info)
            
            RFPDPOST = zip(info['REFERENCE'],info['PARTICULARS'],info['COSTCENTER'],info['AMOUNT'])
            print(info['DUEDATE'],info['AMTINWORDS'],info['TOTALAMOUNT'])
            if form.is_valid():
                instance = form.save(commit=False)
                instance.DEPT = assigned_user.department.department #type:ignore

                PCVTORFPNO_RFP = ""
                RFPBOOL = False
                
                instance.save()
                
                for REFERENCE,PARTICULARS,COSTCENTER,AMOUNT in RFPDPOST:
                    if AMOUNT !="":
                        if RFPBOOL:
                            if not PCVTORFPNO_RFP == REFERENCE:
                                REFERENCE = PCVTORFPNO_RFP

                        if not Location.objects.filter(location=COSTCENTER).exists() and not Department.objects.filter(department=COSTCENTER).exists():
                            data['form_is_valid'] = False
                            data['MSG'] = "Cost Center not existing"
                            csvfile_path_from = directory_path + '/{}'.format(csv_file)
                            destination_folder = destination_folder_path + '/LOGS/ERROR/CSV/'
                            
                            if os.path.exists(csvfile_path_from):
                                shutil.move(csvfile_path_from, destination_folder)
                                print("Folder '{}' successfully moved to '{}'.".format(csvfile_path_from, destination_folder))
                            else:
                                print("Source folder '{}' does not exist.".format(csvfile_path_from))
                            
                            carbon_copy = []
                            typeof = []
                            files = []
                            error_list.append(str(data['MSG']))
                            status_list.append('Error')

                        if COSTCENTER != "N/A":
                            RFPDetails.objects.create(RFP_ID=instance
                            ,REFERENCE=REFERENCE
                            ,PARTICULARS=PARTICULARS
                            ,COST_NAME=COSTCENTER
                            ,AMOUNT=AMOUNT)
                        else:
                            pass

                RFPhead = RFPHeader.objects.get(id=instance.id)
                RFPhead.LOC= assigned_user.location.location #type:ignore
                RFPhead.PREPAREDBY=str(assigned_user.user.first_name) + " " + str(assigned_user.user.last_name)
                RFPhead.DATE=date.today().strftime("%m-%d-%Y")
                RFPhead.USERID=assigned_user.id #type:ignore
                RFPhead.PCVTORFPNO=PCVTORFPNO_RFP
                RFPhead.RFPREPLENISHMENT=RFPBOOL
                RFPhead.location =  assigned_user.location
                RFPhead.department =  assigned_user.department
                RFPhead.save()

                PCVtoRFPNo.objects.create(DEPTSERIESPCVNO= assigned_user.deptseries,PCVTORFPNO=PCVTORFPNO_RFP)
                data['form_is_valid'] = True
                data['MSG'] = "Successfully Created!"
                status_list.append('Created')
                csvfile_path_from = directory_path + '/{}'.format(csv_file)
                destination_folder = destination_folder_path + '/ARCHIVED/' + current_year + '/PROCESSED_CSV/'

                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                else:
                    pass
                
                if os.path.exists(csvfile_path_from):
                    shutil.move(csvfile_path_from, destination_folder)
                else:
                    pass
            else:
                data['form_is_valid'] = False
                data['MSG'] = "but with Invalid Inputs!"

                csvfile_path_from = directory_path + '/{}'.format(csv_file)
                destination_folder = destination_folder_path + '/LOGS/ERROR/'

                if os.path.exists(csvfile_path_from):
                    shutil.move(csvfile_path_from, destination_folder)
                carbon_copy = []
                typeof = []
                files = []
                error_list.append(str(data['MSG']))
                status_list.append('Error')
                mail_automation(assigned_user.user.email, automation_template_html('AUTOMATION',total_amount_list,file_name_list,payee_list,dates_list,due_date_list,status_list,error_list,files,typeof), files,'INFORM: RFP AUTOMATION COMPLETED ' + file_name ,carbon_copy) # type: ignore

        except Exception as e:
            error_list.append(str(e))
            status_list.append('Error')
            carbon_copy = []
            typeof = []
            files = []
           
    carbon_copy = []
    typeof = []
    files = []
    mail_automation(assigned_user.user.email, automation_template_html('AUTOMATION',total_amount_list,file_name_list,payee_list,dates_list,due_date_list,status_list,error_list,files,typeof), files,'INFORM: RFP AUTOMATION COMPLETED GLOBE POSTPAID' ,carbon_copy) # type: ignore
    return JsonResponse(data)
    

@transaction.atomic
def rfp_automate_innove(request,*args,**kwargs):
    token = request.headers.get('Authorization')
    if not token:
        return JsonResponse({'error': 'Token not provided'}), 401

    assigned_user = Profile.objects.get(selectedForAutomation = True)
    info = {}
    data = dict()
    file_name_list = []
    total_amount_list =[]
    payee_list = []
    dates_list = []
    due_date_list = []
    status_list = []
    error_list = []
    location_list = []
    directory_path = r'\\172.18.118.9\d\ADMINISTRATION\INNOVE\EXTRACTED'
    destination_folder_path = r'\\172.18.118.9\d\ADMINISTRATION\INNOVE'
    RFPDPOST = ""
    current_year = str(datetime.now().year)
    all_files = os.listdir(directory_path)
    csv_files = [file for file in all_files if file.lower().endswith('.csv')]
    for csv_file in csv_files:
        file_name = os.path.splitext(os.path.basename(csv_file))[0]
        csvfile_path = directory_path + '/{}'.format(csv_file)

        info['REFERENCE'] = get_reference_values(csvfile_path,'innove') #type:ignore
        info['PARTICULARS'] = get_particulars_values(csvfile_path,'innove') #type:ignore
        info['AMOUNT'] = get_amount_values(csvfile_path,'innove') #type:ignore
        info['COSTCENTER'] = get_cost_center_values(csvfile_path,'innove') #type:ignore
        
        info['DUEDATE'] = str(get_due_date(csvfile_path,'innove')) #type:ignore
        info['AMTINWORDS'] = get_amt_word(info['AMOUNT'],'innove') #type:ignore
        info['TOTALAMOUNT'] = get_total_amt(info['AMOUNT'],'innove') #type:ignore
        info['PAYTO'] = "INNOVE COMMUNICATIONS, INC."
        file_name_list.append(str(file_name))
        total_amount_list.append(str(info['TOTALAMOUNT']))
        payee_list.append(str(info['PAYTO']))
        dates_list.append(str(date.today().strftime("%m-%d-%Y")))
        due_date_list.append(str(info['DUEDATE']))

        try:
            form = Create_RFPForm(info)
            
            RFPDPOST = zip(info['REFERENCE'],info['PARTICULARS'],info['COSTCENTER'],info['AMOUNT'])
            print(len(info['REFERENCE']),len(info['PARTICULARS']),len(info['COSTCENTER']),len(info['AMOUNT']),)
            for COSTCENTERS in info['COSTCENTER']:
                if not Location.objects.filter(location=COSTCENTERS).exists() and not Department.objects.filter(department=COSTCENTERS).exists():
                    data['form_is_valid'] = False
                    location_list.append("{}".format(COSTCENTERS))

            if form.is_valid():
                instance = form.save(commit=False)
                instance.DEPT = assigned_user.department.department #type:ignore

                PCVTORFPNO_RFP = ""
                RFPBOOL = False
                
                instance.save()
                
                for REFERENCE,PARTICULARS,COSTCENTER,AMOUNT in RFPDPOST:
                    if AMOUNT !="":
                        if RFPBOOL:
                            if not PCVTORFPNO_RFP == REFERENCE:
                                REFERENCE = PCVTORFPNO_RFP

                        if Location.objects.filter(location=COSTCENTER).exists() or Department.objects.filter(department=COSTCENTER).exists():
                            if COSTCENTER != "N/A":
                                RFPDetails.objects.create(RFP_ID=instance
                                ,REFERENCE=REFERENCE
                                ,PARTICULARS=PARTICULARS
                                ,COST_NAME=COSTCENTER
                                ,AMOUNT=AMOUNT)
                            else:
                                pass
                        else:
                            data['form_is_valid'] = False
                            location_list.append("Location not Existing: {}".format(COSTCENTER))

                RFPhead = RFPHeader.objects.get(id=instance.id)
                RFPhead.LOC= assigned_user.location.location #type:ignore
                RFPhead.PREPAREDBY=str(assigned_user.user.first_name) + " " + str(assigned_user.user.last_name)
                RFPhead.DATE=date.today().strftime("%m-%d-%Y")
                RFPhead.USERID=assigned_user.id #type:ignore
                RFPhead.PCVTORFPNO=PCVTORFPNO_RFP
                RFPhead.RFPREPLENISHMENT=RFPBOOL
                RFPhead.location =  assigned_user.location
                RFPhead.department =  assigned_user.department
                RFPhead.save()

                PCVtoRFPNo.objects.create(DEPTSERIESPCVNO= assigned_user.deptseries,PCVTORFPNO=PCVTORFPNO_RFP)
                data['form_is_valid'] = True
                data['MSG'] = "Successfully Created!"
                status_list.append('Created')
                csvfile_path_from = directory_path + '/{}'.format(csv_file)
                destination_folder = destination_folder_path + '/ARCHIVED/' + current_year + '/PROCESSED_CSV/'

                if not os.path.exists(destination_folder):
                    os.makedirs(destination_folder)
                else:
                    pass
                
                if os.path.exists(csvfile_path_from):
                    shutil.move(csvfile_path_from, destination_folder)
                else:
                    pass
            else:
                data['form_is_valid'] = False
                data['MSG'] = "but with Invalid Inputs!"

                csvfile_path_from = directory_path + '/{}'.format(csv_file)
                destination_folder = destination_folder_path + '/LOGS/ERROR/CSV/'

                if os.path.exists(csvfile_path_from):
                    shutil.move(csvfile_path_from, destination_folder)
                carbon_copy = []
                typeof = []
                files = []
                error_list.append(str(data['MSG']))
                status_list.append('Error')
                    
        
        except Exception as e:
            carbon_copy = []
            typeof = []
            files = []
            error_list.append(str(e))
            status_list.append('Error')
            mail_automation(assigned_user.user.email, automation_template_html('AUTOMATION',total_amount_list,file_name_list,payee_list,dates_list,due_date_list,status_list,error_list,files,typeof), files,'INFORM: RFP AUTOMATION COMPLETED ' + file_name,carbon_copy) # type: ignore
    carbon_copy = []
    typeof = []
    files = []
    if len(location_list) != 0:
        error_list.append(str(location_list))
    mail_automation(assigned_user.user.email, automation_template_html('AUTOMATION',total_amount_list,file_name_list,payee_list,dates_list,due_date_list,status_list,error_list,files,typeof), files,'INFORM: RFP AUTOMATION COMPLETED INNOVE',carbon_copy) # type: ignore

    return JsonResponse(data)

@login_required
def rfp_assign_user_for_automation(request):
    data = dict()
    data["form_is_valid"] = True
    form = SelectUserForAutomation(user=request.user)
    context = {'form': form}
    data['html_form'] = render_to_string('rfp/includes/part_rfp_select_user.html',
        context,
        request=request,
    )
    return JsonResponse(data)

@login_required
def rfp_assign(request):
    data = dict()
    try:
        if request.method == 'POST':
            form = SelectUserForAutomation(request.POST,user=request.user)
            if form.is_valid():
                data['is_valid'] = True
                try:
                    current_profile = Profile.objects.get(selectedForAutomation = True)
                    current_profile.selectedForAutomation = False
                    current_profile.save()
                except:
                    pass

                selected_user = form.cleaned_data.get('selectedForAutomation')
                profile = Profile.objects.get(id = selected_user)
                profile.selectedForAutomation = True
                profile.save()
                
            else:
                data['form_is_valid'] = False
        else:
            form = SelectUserForAutomation(user=request.user)
        data['html_form'] = render_to_string('rfp/includes/part_rfp_select_user.html',
            {'form':form},
            request=request
        )
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()

    return JsonResponse(data)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def rental_maintenance(request):
    template_name = "rental/rental.html"
    data = dict()
    context = {}
    context['title'] = 'Rental Automation'
    data['html_form'] = render_to_string(template_name,context,request=request)

    return JsonResponse(data)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def rental_table(request):
    template_name = "rental/rental_table.html"
    data = dict()
    context = {}

    context['title'] = 'Rental Automation Table '
    data['html_form'] = render_to_string(template_name,context,request=request)

    return JsonResponse(data)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def get_rental(request):
    
    return JsonResponse(request)

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def add_rental(request):
    
    return JsonResponse(request)

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def update_rental(request):
    
    return JsonResponse(request)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def get_rental_year(request):
    rental_year = Rental_Scoped_Date.objects.all()
    serializer = RentalYearSerializer(rental_year,many=True)
    return Response({'year': serializer.data})

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def render_postpaid_table(request, *args, **kwargs):
    template_name = "postpaid/postpaid_table.html"

    data = dict()
    context = {}
    locations = []
    departments = []

    try:
        locations = Location.objects.filter(Q(Status = "O",Store_Ho = True) | Q(Status='A',Main_Ho = True) | Q(Status='A',Sub_Ho = True)).values('id','location')
        departments = Department.objects.filter(~Q(Status='C')).values('id','department')
    except:
        pass

    location_list = []
    for pv in locations:
        location_data = {}
        location_data['id'] = pv['id']
        location_data['name'] = pv['location']
        location_list.append(location_data)
    data['locations'] = location_list

    department_list = []
    for pv in departments:
        department_data = {}
        department_data['id'] = pv['id']
        department_data['name'] = pv['department']
        department_list.append(department_data)
    data['locations'].extend(department_list)

    context['title'] = 'Postpaid Table'
    data['html_form'] = render_to_string(template_name,context,request=request)
    return JsonResponse(data)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def render_innove_table(request, *args, **kwargs):
    template_name = "innove/innove_table.html"

    data = dict()
    context = {}
    location = []
    department = []

    try:
        location = Location.objects.filter(Q(Status = "O",Store_Ho = True)).values('id','location')
        department = Department.objects.filter(~Q(Status='C')).values('id','department')
    except:
        pass

    location_list = []
    for pv in location:
        location_data = {}
        location_data['id'] = pv['id']
        location_data['name'] = pv['location']
        location_list.append(location_data)
    data['location'] = location_list

    department_list = []
    for pv in department:
        department_data = {}
        department_data['id'] = pv['id']
        department_data['name'] = pv['department']
        department_list.append(department_data)
    data['location'].extend(department_list)
    
    context['title'] = 'Innove Table'
    data['html_form'] = render_to_string(template_name,context,request=request)
    return JsonResponse(data)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def get_postpaid_table(request,**kwargs):
    data = {}
    context = {}
    datatables = request.GET
    results = []
    start_page = datatables.get('start')
    length = (datatables.get('length'))
    search = datatables.get('search[value]')
    draw = int(datatables.get('draw'))

    postpaid_master_list_obj = PostpaidMasterList.objects.filter(status='A')

    records_total = len(postpaid_master_list_obj)
    records_filtered = records_total
    serializer = PostpaidMasterListSerializer(postpaid_master_list_obj,many=True)

    lists_data = paginate_page(length,start_page,serializer,records_total) # type: ignore

    for data in lists_data:
        results.append(make_row_postpaid_master_list(data)) # type: ignore

    context = { 'data': results, 'recordsTotal': records_total , 'recordsFiltered': records_filtered, 'draw': draw }
    return JsonResponse(context)

@api_view(['GET'])
@permission_classes([IsAuthenticated]) 
def get_innove_table(request,**kwargs):
    data = {}
    context = {}
    datatables = request.GET
    results = []
    start_page = datatables.get('start')
    length = (datatables.get('length'))
    search = datatables.get('search[value]')
    draw = int(datatables.get('draw'))

    innove_master_list_obj = InnoveMasterList.objects.filter(status='A')

    records_total = len(innove_master_list_obj)
    records_filtered = records_total
    serializer = InnoveMasterListSerializer(innove_master_list_obj,many=True)

    lists_data = paginate_page(length,start_page,serializer,records_total) # type: ignore

    for data in lists_data:
        results.append(make_row_innove_master_list(data)) # type: ignore

    context = { 'data': results, 'recordsTotal': records_total , 'recordsFiltered': records_filtered, 'draw': draw }
    return JsonResponse(context)

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def create_postpaid_account(request,**kwargs):
    serializer = CreatePostPaidAccountSerializer(data = request.data)

    if serializer.is_valid():
        PostpaidMasterList.objects.create(
            account_number = request.data['account_number'],
            name = request.data['name'],
            cost_center = request.data['cost_center'],
            status = 'A'
        )
        return JsonResponse({'message': 'success'})
    
    return JsonResponse({'message': 'Failed'})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def update_postpaid_account(request,**kwargs):
    postpaid_account_obj = PostpaidMasterList.objects.get(id = request.data['id'])
    serializer = CreatePostPaidAccountSerializer(postpaid_account_obj, data = request.data, partial=True)

    if serializer.is_valid():
        serializer.save()
        return JsonResponse({'message': 'success'})
    
    return JsonResponse({'message': 'Failed'})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def delete_postpaid_account(request,**kwargs):
    postpaid_account_obj = PostpaidMasterList.objects.get(id = request.data['id'])
    postpaid_account_obj.status = 'IA'
    postpaid_account_obj.save()

    return JsonResponse({'message': 'success'})


@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def delete_innove_account(request,**kwargs):
    innove_account_obj = InnoveMasterList.objects.get(id = request.data['id'])
    innove_account_obj.status = 'IA'
    innove_account_obj.save()

    return JsonResponse({'message': 'success'})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def update_innove_account(request,**kwargs):
    innove_account_obj = InnoveMasterList.objects.get(id = request.data['id'])
    serializer = CreateInnoveAccountSerializer(innove_account_obj, data = request.data, partial=True)

    if serializer.is_valid():
        serializer.save()
        return JsonResponse({'message': 'success'})
    
    return JsonResponse({'message': 'Failed'})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def create_innove_account(request,**kwargs):
    serializer = CreateInnoveAccountSerializer(data = request.data)

    if serializer.is_valid():
        InnoveMasterList.objects.create(
            account_number = request.data['account_number'],
            type = request.data['type'],
            cost_center = request.data['cost_center'],
            status = 'A'
        )
        return JsonResponse({'message': 'success'})
    
    return JsonResponse({'message': 'Failed'})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def upload_postpaid_csv(request,**kwargs):
    start_row = 1
    remarks = []
    existing_record = None
    df = pd.read_csv(request.data['csv'], encoding='latin-1')
    for index, row in df.iloc[start_row:].iterrows():
        
        row_dict = row.to_dict()

        row_dict['account_number'] = str(int(row_dict['account_number']))
        try:
            existing_record = PostpaidMasterList.objects.get(account_number=row_dict['account_number'])
        except:
            pass
        existing_location = Location.objects.filter(location=row_dict['cost_center']).exists()
        existing_department = Department.objects.filter(department=row_dict['cost_center']).exists()
        # If the record doesn't exist, create and save it
        if not existing_record and (existing_location or existing_department):
            innove_instance = PostpaidMasterList(**row_dict)
            innove_instance.save()

        elif existing_record.name != row_dict['name'] or existing_record.cost_center != row_dict['cost_center']: #type:ignore
            if row_dict['name'] != "":
                existing_record.name = row_dict['name'] #type:ignore
                existing_record.save() #type:ignore
                remarks.append("Record with Account Number {} successfully updated Name.<br>".format(row_dict['account_number']))
            else:
                remarks.append("Record with Account Number {} has name "".<br>".format(row_dict['account_number']))

            if existing_record.cost_center != row_dict['cost_center'] and (existing_location or existing_department): #type:ignore
                existing_record.cost_center = row_dict['cost_center'] #type:ignore
                existing_record.save() #type:ignore
                remarks.append("Record with Account Number {} successfully updated Cost Center.<br>".format(row_dict['account_number']))
            else:
                remarks.append("Record with Account Number {} location not updated.<br>".format(row_dict['account_number']))

        else:
            remarks.append("Record with Account Number {} already exists. Skipping upload.<br>".format(row_dict['account_number']))
        
    return JsonResponse({'message': 'success','remarks': remarks})

@api_view(['POST'])
@permission_classes([IsAuthenticated]) 
def upload_innove_csv(request,**kwargs):
    start_row = 1
    existing_record = None
    remarks = []
    df = pd.read_csv(request.data['csv'], encoding='latin-1')
    for index, row in df.iloc[start_row:].iterrows():
        row_dict = row.to_dict()
        row_dict['account_number'] = str(int(row_dict['account_number']))
        try:
            existing_record = InnoveMasterList.objects.get(account_number=row_dict['account_number'])
        except:
            pass
        existing_location = Location.objects.filter(location=row_dict['cost_center']).exists()
        existing_department = Department.objects.filter(department=row_dict['cost_center']).exists()

        # If the record doesn't exist, create and save it
        if not existing_record and (existing_location or existing_department):
            innove_instance = InnoveMasterList(**row_dict)
            innove_instance.save()

        elif existing_record.type != row_dict['type'] or existing_record.cost_center != row_dict['cost_center']: #type:ignore
            if row_dict['type'] != "":
                existing_record.type = row_dict['type'] #type:ignore
                existing_record.save() #type:ignore
                remarks.append("Record with Account Number {} successfully updated Type.<br>".format(row_dict['account_number']))
            else:
                remarks.append("Record with Account Number {} has type "".<br>".format(row_dict['account_number']))

            if existing_record.cost_center != row_dict['cost_center'] and (existing_location or existing_department): #type:ignore
                existing_record.cost_center = row_dict['cost_center'] #type:ignore
                existing_record.save() #type:ignore
                remarks.append("Record with Account Number {} successfully updated Cost Center.<br>".format(row_dict['account_number']))
            else:
                remarks.append("Record with Account Number {} location not updated.<br>".format(row_dict['account_number']))

        else:
            remarks.append("Record with Account Number {} already exists. Skipping upload.<br>".format(row_dict['account_number']))
    
    return JsonResponse({'message': 'success','remarks': remarks})