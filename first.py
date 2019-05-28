import time    
from datetime import datetime
import yaml
import csv
import string
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import NameObject, TextStringObject, IndirectObject

errors = ""
unique_program_names = {} 
def format_name(unformatted_name):
    global errors
    if unformatted_name == "": 
        errors += "name is empty"
    if "," not in unformatted_name: 
        errors += "name '%s' has no comma\n" % (unformatted_name)
        formatted_name = string.capwords(unformatted_name.strip())
    else:
        name_array = unformatted_name.split(',')
        if len(name_array) != 2:
            errors += "client name has %d commas!\n" % len(name_array)
        formatted_name = string.capwords(name_array[1].strip()) + " " + string.capwords(name_array[0].strip())
    return formatted_name

with open("config.txt", 'r') as config_data:
    config = yaml.safe_load(config_data)
template_filename = "input/%s" % config['template_filename']
template_reader = PdfFileReader(open(template_filename, 'rb'))



if config['run_data_refresh']:
    class Scope:
        # :debug find fileds on page
        if False:
            for field in template_reader.getFields():
                print(field)
            exit()

        print("------------- CSV DATA --------------")
        csv_data = []
        with open("input/%s" % config["raw_data_filename"], 'r') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            for csv_row in csv_reader:
                csv_data.append(csv_row)
        row_id = 0;
        value_ndx = 0;
        value_array = []
        for row in csv_data:
            row_id += 1
            values = {}
            values['row_id'] = row_id
            values['errors'] = ""
            global errors
            errors = ""
            values['program_count'] = 0
            values['program'] = ""
            values['program_fap'] = ""
            values['program_medicaid'] = ""
            values['program_medicare'] = ""
            values['program_cdc'] = ""
            values['program_cash'] = ""
            values["page_6_hide"] = False
            values["page_7_hide"] = False
            values["page_8_hide"] = False
            values["page_9_hide"] = False
            values["page_10_hide"] = False
            for col in row:
                values[col] = "%s" % row[col]
                ignore_length_validation = ["office", "program", "office_address"]
                if len(values[col]) > 24 and col not in ignore_length_validation:
                    errors += "Field is size %d : '%s':'%s'\n" %(len(values[col]), col, values[col])

            values['program_raw'] = values['program']
            # don't use row beyond this point

            # formatting found fields
            values['client_name'] = format_name(values['client_name']) 
            values['full_name'] = format_name(values['full_name']) 
        
            for i in range(1, 4):
                string_formatting = 'member_%d_legal_name'
                index_name = string_formatting  % i
                if values[index_name] != "":
                    values[index_name] = format_name(values[index_name])
                #print("'%s'" %values[index_name])

            for i in range(1, 6):
                string_formatting = 'income_%d_recipient'
                index_name = string_formatting  % i
                if values[index_name] != "":
                    values[index_name] = format_name(values[index_name])
                #print("'%s'" %values[index_name])

# income_2_recipient
# income_6_recipient



            if values['office'] == "Madison Heights Field Office":
                values['office_address'] = "Madison Heights MDHHS Office\n30755 MONTPELIER\nMADISON HEIGHTS, MI 48071"
            elif values['office'] == "Monroe Field Office":
                values['office_address'] = "Monroe County MDHHS Office\n903 S. Telegraph, Ste. A\nMonroe, MI 48161"
            else:
                assert false, "office not found: '%s'" % values['office']


            # :generated fields
            # :program
            if values['program_raw'] in unique_program_names: 
                unique_program_names[values['program_raw']] += 1
            else:
                unique_program_names[values['program_raw']] =  1

            values['program_fap'] = "FAP" in values['program_raw']
            values['program_medicaid'] = "Medicaid" in values['program_raw']
            values['program_medicare'] = "Medicare Cost Share" in values['program_raw']
            values['program_cdc'] = "CDC" in values['program_raw']
            values['program_cash'] = "Cash" in values['program_raw']

    
            if values['program_fap'] :
                values['program_count'] += 1
            if values['program_medicaid'] :
                values['program_count'] += 1
            if values['program_medicare'] :
                values['program_count'] += 1
            if values['program_cdc'] :
                values['program_count'] += 1
            if values['program_cash'] :
                values['program_count'] += 1

            values['program'] = ""
            if values['program_count'] <= 2:
                values['program'] = "\n"

            # regenerate program
            if values['program_fap'] :
                values['program'] += "Food Assistance Program\n"
            if values['program_medicaid'] :
                values['program'] += "Medicaid\n"
            if values['program_medicare'] :
                values['program'] += "Medicare Cost Share\n"
            if values['program_cdc'] :
                values['program'] += "Child Development + Care\n"
            if values['program_cash'] :
                values['program'] += "Cash\n"

            assert values['program_count'] <= 3, "program count should never be above 3 row_id %d" % values['row_id']

            values['program'].rstrip('\n')


            if values['program_fap'] or values['program_cash']:
                assert values['interview_type'] is not 'None'
                if values['interview_date'].strip() == '':
                    errors += "interview date_skipped because it is empty.\n"
            else:
                assert 'None' in values['interview_type'], "row_id: %d Interview type is not %s but it should be None for program %s" % (values['row_id'], values['interview_type'], values['program_raw'])


            values["page_count"] = 11;

            if not values['program_fap']:
                values["page_6_hide"]   = True
                values["page_count"] -= 1

            if (not values['program_medicaid'] and not values['program_medicare']):
                values["page_7_hide"]   = True
                values["page_8_hide"]   = True
                values["page_count"] -= 2

            if not values['program_cdc']:
                values["page_9_hide"]   = True
                values["page_count"] -= 1

            if not values['program_cash']:
                values["page_10_hide"]  = True
                values["page_count"] -= 1


            # assert has at least one program
            assert values['program_fap'] or values['program_medicaid'] or values['program_medicare'] or values['program_cdc'] or values['program_cash'], "No programs found in program column '%s'" % values['program_raw']



            # :name_and_address
            if values["mailing_street"] == "":
                # is apartment lot really required?
                values['name_and_address'] = "%s\n%s %s\n%s, %s %s" % (values['client_name'], values['street'], values['apartment_lot'], values['city'], values['state'], values['zipcode'])
            else:
                # @Todo: Remove astrisk
                values['name_and_address'] = "%s\n%s %s\n%s, %s %s" % (values['client_name'], values['mailing_street'], values['mailing_apartment_lot'], values['mailing_city'], values['state'], values['mailing_zipcode'])

            # :due_date
            values['due_date_full'] = ""
            if values['due_date'] != '':
                date = datetime.strptime(values['due_date'], '%m/%d/%Y')
                if date.year == 1900:
                    date.replace(year=2019)
                formatted_date = date.strftime('%b %d, 2019')
                values['due_date_full'] = formatted_date
                values['due_date'] = "%s/%s/%s" % (date.strftime('%m').lstrip('0'), date.strftime('%d').lstrip('0'), date.strftime('%y'))
            else:
                errors += "due date skipped.\n"


            if values['interview_date'] != '':
                date = datetime.strptime(values['interview_date'], '%m/%d/%Y')
                if date.year == 1900:
                    date.replace(year=2019)
                values['interview_date'] = "%s/%s/%s" % (date.strftime('%m').lstrip('0'), date.strftime('%d').lstrip('0'), date.strftime('%y'))

            else:
                errors += "interview date skipped.\n"

            values['interview_date_and_time'] = "N/A"
            if values['interview_type'] != "None":
                values['interview_date_and_time'] = "%s at %s" % (values['interview_date'], values['interview_time'])
            else:
                if values["interview_date"] != "" or values['interview_time'] != "": 
                    errors += "Interview date and time should be blank: %s date:'%s' time:'%s' type:'%s'\n" % (values['row_id'], values['interview_date'], values['interview_time'], values['interview_type'])

            values['fap_text'] = ""
            if values['program_fap']:
                values['fap_text'] = "Your Food Assistance Program (FAP) benefits will end on 7/31/2019. You must submit your redetermination form or filing form by 7/15/2019 in order to receive uninterrupted FAP benefits."
            # :tax

            count_of_people_who_paid_taxes = 0;
            last_person_who_filed_taxes = ""
            for i in range(1,8):
                if values["member_%d_taxes" % i] == "Yes":
                    count_of_people_who_paid_taxes += 1 
                    last_person_who_filed_taxes = values['member_%d_legal_name' % i]

            only_one_household_member_filed_taxes = count_of_people_who_paid_taxes == 1
            no_household_members_file_taxes =  count_of_people_who_paid_taxes == 0


            values['tax_name'] = '' 
            values['tax_check_no'] = 'No' 
            values['tax_check_yes'] = 'No' 
            # Don't set these if we're not on the healthcare coverage page
            if not values['page_8_hide']:
                if only_one_household_member_filed_taxes:
                    values['tax_check_yes'] = 'Yes'
                    values['tax_check_no'] = 'No'
                    values['tax_name'] = format_name(last_person_who_filed_taxes)
                if no_household_members_file_taxes:
                    values['tax_check_yes'] = 'No'
                    values['tax_check_no'] = 'Yes'



            flag = ''
            #if values['tax_check_yes'] == "Yes":
            if values['mailing_street'] != "":
                flag = "X_"

            values['pdf_filename'] = "%03d_%s%02d_%s_%s.pdf" % (values['row_id'], flag, values['page_count'], values['program_raw'], values['case_number'])

            if errors != "":
                print("errors: %s" %errors)
            values['errors'] = errors
            value_array.append(values)


        # :writer
        with open('output/data.csv', 'w') as csv_file:
            csv_writer = csv.writer(csv_file, delimiter=',', lineterminator='\n')
            created_headers = []
            for col in value_array[0]:
                created_headers.append(col)
            csv_writer.writerow(created_headers)

            for row in value_array: 
                created_row = []
                for col in row:
                    created_row.append(row[col])
                csv_writer.writerow(created_row)
    print ("---------------------- data step complete ----------------------")

    print ("---------------------- unique program names ----------------------")

    print ("unique program names")
    for name in unique_program_names:
        print("%-80s %d" %(name, unique_program_names[name]))

    print ("---------------------- unique program names ----------------------")
else:
    print("> Skipping Data Refresh")



if config['run_pdf_generation']:
    class Scope:
        print ("---------------------- begin export ----------------------")

        print(time.strftime('%Y-%m-%d %H:%M:%S'))
        csv_file = open("output/data.csv")
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        imported_value_array = []
        for row in csv_reader:
            imported_value_array.append(row)

        print("rows %d"% len(imported_value_array))
        # :logic
        logic_fields = ["program_raw", "program_count", "errors", "page_count", "form", "row_id", "pdf_filename", "page_6_hide", "page_7_hide", "page_8_hide", "page_9_hide", "page_10_hide", "program_fap", "program_medicaid", "program_medicare", "program_cdc", "program_cash", "interview_date", "interview_time", "street", "apartment_lot", "city", "zipcode", "state", "mailing_street", "mailing_apartment_lot", "mailing_city", "mailing_zipcode"]
        for i in range(1,9):
            logic_fields.append("member_%d_taxes" % i)
        missing_template_fields = {}
        for ndx, values in enumerate(imported_value_array):
            if not config['generate_all_files']:
                if ndx == 15: break
            print("writing %s" % values['row_id'])
            writer = PdfFileWriter()
            fields_filled = []
            # scan pages for annot fields
            for pageNum in range(template_reader.numPages):
                # Skip pages based on program
                # :pages
                if pageNum == 6 and values["page_6_hide"] == 'True':
                    continue
                if pageNum == 7 and values["page_7_hide"] == 'True':
                    continue
                if pageNum == 8 and values["page_8_hide"] == 'True':
                    continue
                if pageNum == 9 and values["page_9_hide"] == 'True':
                    continue
                if pageNum == 10 and values["page_10_hide"] == 'True':
                    continue
                page = template_reader.getPage(pageNum)
                writer.addPage(page)
                if "/Annots" in page:
                    # writer.updatePageFormFieldValues(page, values)
                    for j in range(0, len(page['/Annots'])):
                        annot_child = page['/Annots'][j].getObject()
                        for field in values:
                            if annot_child.get('/Parent'):
                                annot_child_parent = annot_child.get('/Parent').getObject() 
                                if annot_child_parent.get('/T') == field:
                                    annot_child_parent.update({
                                        NameObject("/V"): TextStringObject(values[field]),
                                        })
                                    fields_filled.append(field)
                            if annot_child.get('/T') == field:
                                annot_child.update({
                                    NameObject("/V"): TextStringObject(values[field])
                                    })
                                fields_filled.append(field)

            # check to see if everything was filled
            # some values are never in the pdf (logic only)
            print("%s" % values['pdf_filename']);
            # other values are sometimes not available to be filled
            fields_not_visible = []
            # @Todo: Testing - set some invalid data and see if it picks up on it.
            if values["page_8_hide"]:
                fields_not_visible.append("tax_check_yes")
                fields_not_visible.append("tax_check_no")
                fields_not_visible.append("tax_name")
            else:
                if values["tax_check_no"]:
                    fields_not_visible.append("tax_name")

            for k in values:
                if k not in fields_filled and k not in logic_fields and k not in fields_not_visible:
                    if k not in missing_template_fields:
                        missing_template_fields[k] = 1
                    else:
                        missing_template_fields[k] += 1
                    #print("template field not found: %s" % k)

            print("pdf_filename '%s'" % values['pdf_filename'])
            writer.write(open("output/%s" % values['pdf_filename'], "wb"))
        if len(missing_template_fields) > 0:    
            print("missing template fields:")
        for field in missing_template_fields:
            print(field)
else:
    print("> Skipping PDF Generation")

if config['run_data_validation']:
    print ("---------------------- begin data validation----------------------")

    print(time.strftime('%Y-%m-%d %H:%M:%S'))
    csv_file = open("output/data.csv")
    csv_reader = csv.DictReader(csv_file, delimiter=',')
    data_array = []
    for csv_reader_row in csv_reader:
        data_array.append(csv_reader_row)
    
    print("Validating Page Numbers")
    for ndx, data_dict in data_array:
        if not config['generate_all_files']:
            if ndx == 15: break
        # Open up PDF
        print("Validating %s" % data_dict["pdf_filename"])
        with open("output/%s" % data_dict["pdf_filename"], 'rb') as generated_pdf:
            gen_pdf_reader = PdfFileReader(generated_pdf)
            assert int(gen_pdf_reader.getNumPages()) == int(data_dict['page_count']), "Page Number Issue: pdf has: '%s'  data has: '%s'" %(gen_pdf_reader.getNumPages(), data_dict['page_count'])
    print("valid")


    

