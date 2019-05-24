import time    
from datetime import datetime
import yaml
import csv
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import NameObject, TextStringObject, IndirectObject

with open("config.yml", 'r') as config_data:
    config = yaml.safe_load(config_data)
template_filename = "input/%s" % config['template_filename']
template_reader = PdfFileReader(open(template_filename, 'rb'))

class Scope:
    # :debug find fileds on page
    if False:
        print(template_reader.getPage(8)['/Annots'].getObject())
        for annot in template_reader.getPage(0)['/Annots'].getObject():
            print('>')
            for k1, v1 in annot.getObject().items():
                print("  %s"%k1)
                if k1 == '/Parent':
                    for k2, v2 in v1.getObject().items():
                        print("    %s" %k2)
                        if k2 == '/Kids':
                            print("      %s" %k2)
                            print("        %s" %v2.getObject())
                            if isinstance(v2.getObject(), list):
                                for o in v2.getObject():
                                    print(repr(o.getObject()))

                        else:
                            print("        %s" %v2.getObject())
                else :
                    print("    %s" %v1)
        exit()

    if False:
        print(repr(template_reader.getFields()))
        exit()
    #print(repr(template_reader.getFormTextFields()))
    template_fields = template_reader.getFormTextFields()
    vales = {}

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

        required_not_blank = ['client_name', 'interview_date', 'interview_time']
        for col_name in required_not_blank:
            if row[col_name] == '':
                # @todo: log this error
                print("column %s is empty in application %s" % (col_name, row['client_name']))

        # :generated fields
        values['program_fap'] = "FAP" in row['program']
        values['program_medicaid'] = "Medicaid" in row['program']
        values['program_medicare'] = "Medicare Cost Share" in row['program']
        values['program_cdc'] = "CDC" in row['program']
        # @Todo: Add program cash criteria
        values['program_cash'] = False

        values["page_6_hide"] = False
        if not values['program_fap']:
            values["page_6_hide"]   = True

        values["page_7_hide"] = False
        values["page_8_hide"] = False
        if (not values['program_medicaid'] and not values['program_medicare']):
            values["page_7_hide"]   = True
            values["page_8_hide"]   = True

        values["page_9_hide"] = False
        if not values['program_cdc']:
            values["page_9_hide"]   = True

        values["page_10_hide"] = False
        if not values['program_cash']:
            values["page_10_hide"]  = True


        # assert has at least one program
        assert values['program_fap'] or values['program_medicaid'] or values['program_medicare'] or values['program_cdc'] or values['program_cash'], "No programs found in program column '%s'" % row['program']



        # :name_and_address
        # @Todo: parse mailing_address
        if row["mailing_street"] == "":
            # is apartment lot really required?
            values['name_and_address'] = "%s\n%s %s\n%s, %s %s" % (row['client_name'], row['street'], row['apartment_lot'], row['city'], row['state'], row['zipcode'])
            #print("%s\n" % values['name_and_address'])
        else:
            # @Todo: Remove astrisk
            values['name_and_address'] = "%s\n%s %s\n%s, %s %s" % (row['client_name'], row['mailing_street'], row['mailing_apartment_lot'], row['mailing_city'], row['state'], row['mailing_zipcode'])
            #print("%s\n" %values['name_and_address'])

        # :due_date
        if row['due_date'] != '':
            date = datetime.strptime(row['due_date'], '%m/%d/%Y')
            if date.year == 1900:
                date.replace(year=2019)
            formatted_date = date.strftime('%a %b %d, 2019')
            row['due_date'] = formatted_date
            row['due_date_full'] = formatted_date
        else:
            print("due date skipped")

        #print("due date %s" % row['due_date'])

        if row['interview_date'] != '':
            date = datetime.strptime(row['interview_date'], '%m/%d/%Y')
            if date.year == 1900:
                date.replace(year=2019)
            formatted_date = date.strftime('%a %b %d')
            row['interview_date'] = formatted_date
        else:
            print("interview date skipped")



        # :interview_date
        # @Todo: format interview date
        values['interview_date'] = row['interview_date']
        values['interview_time'] = row['interview_time']
        # interview_date_and_time
        # interview_date_and_time = interview date + “ at ” + interview_time


        # @Concern
        values['interview_date_and_time'] = ""
        if row['interview_type'] != "None":
            values['interview_date_and_time'] = "%s at %s" % (values['interview_date'], values['interview_time'])
        else:
            print("Missing date or time for interview")
        #print("%s\n" %values['interview_date_and_time'])

        # @Urgent: This is a hack that means that values is not being cleared everytime!!
        if values['program_fap']:
            values['fap_text'] = "Your Food Assistance Program (FAP) benefits will end on 7/31/2019. You must submit your redetermination form or filing form by 7/15/2019 in order to receive uninterrupted FAP benefits."

        # :tax

        # _Tax Filers_
        # if ONLY 1 household member indicates yes they file taxes, 
        # check Yes box. 
        # if member_x_taxes = “Yes”
        # put member_x_legal_name in the tax_name.
        # // I wasn’t sure how to do this, but I was thinking it’d be cool if we could get the name of whoever indicated they were filing taxes to fill in tax_name. The x just goes 1-8 -- there aren’t actually member_x variables. 
        # if ALL household members indicate no they don’t file taxes, 
        # check No box. 
        # else (if any other combinations of answers exist), 
        # 
        # don’t check anything. 


        # :tax

        count_of_people_who_paid_taxes = 0;
        last_person_who_filed_taxes = ""
        for i in range(1,6):
            if row["member_%d_taxes" % i] == "Yes":
                count_of_people_who_paid_taxes += 1 
                last_person_who_filed_taxes = row['member_%d_legal_name' % i]

        only_one_household_member_filed_taxes = count_of_people_who_paid_taxes == 1
        no_household_members_file_taxes =  count_of_people_who_paid_taxes == 0

        # @Todo: Thisis a MAJOR concern. Why isn't values cleared everytime?
        values['tax_name'] = None
        values['tax_check_no'] = None
        values['tax_check_yes'] = None
        if only_one_household_member_filed_taxes:
            values['tax_check_yes'] = 'Yes'
            values['tax_check_no'] = 'No'
            values['tax_name'] = last_person_who_filed_taxes

        if no_household_members_file_taxes:
            values['tax_check_yes'] = 'No'
            values['tax_check_no'] = 'Yes'


        # /generated fields

        # @MajorConcern: this has the ability to overwrite values, causing confusion. 
        # I wonder if I should not do automatic copying
        for f in template_fields:
            if f in row:
                if row[f] != "":
                    values[f] = "%s" % row[f]
                else: 
                    values[f] = "*empty*"
                #print("csv_row %s" % f)
                # print("%s %s" % (f, row[f]))

        values['pdf_filename'] = "%03d_%s_%s.pdf" % (values['row_id'], values['client_name'], values['case_number'])


        value_array.append(values)


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




print ("---------------------- begin export ----------------------")
# This is completely separated from the above code

print(time.strftime('%Y-%m-%d %H:%M:%S'))
csv_file = open("output/data.csv")
csv_reader = csv.DictReader(csv_file)
imported_value_array = []
for row in csv_reader:
    print(row)
    imported_value_array.append(row)

print("rows %d"% len(imported_value_array))
for values in imported_value_array:
    print("writing %s" % values['row_id'])
    writer = PdfFileWriter()
    fields_filled = []
    # scan pages for annot fields
    for pageNum in range(template_reader.numPages):
        # Skip pages based on program
        # :pages
        if pageNum == 6 and values["page_6_hide"]:
            continue
        if pageNum == 7 and values["page_7_hide"]:
            continue
        if pageNum == 8 and values["page_8_hide"]:
            continue
        if pageNum == 9 and values["page_9_hide"]:
            continue
        if pageNum == 10 and values["page_10_hide"]:
            continue
        page = template_reader.getPage(pageNum)
        writer.addPage(page)
        if "/Annots" in page:
            # writer.updatePageFormFieldValues(page, values)
            for j in range(0, len(page['/Annots'])):
                annot_child = page['/Annots'][j].getObject()
                # clear object
                #                    if type(annot_child) != NoneType:
                #                        annot_child.update({NameObject("/V"): ""})
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
    logic_fields = ["row_id", "pdf_filename", "page_6_hide", "page_7_hide", "page_8_hide", "page_9_hide", "page_10_hide", "program_fap", "program_medicaid", "program_medicare", "program_cdc", "program_cash", "interview_date", "interview_time"]
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
            print("automatic value never entered into template: %s" % k)

    writer.write(open("output/%s" % values['pdf_filename'], "wb"))



