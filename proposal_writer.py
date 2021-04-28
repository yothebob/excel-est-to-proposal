from docx import Document
from datetime import date
from docx.shared import Pt, Inches,RGBColor
import sys
from tkinter import *
from tkinter.filedialog import asksaveasfilename ,askopenfilename
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import re



print('excel to py file loaded...')


sys.path.append('C:/Users/Owner/Desktop/Estimating model 1.0.7.9')

file = ''
def set_file(_file):
    global file
    file = _file
    print(file)

estimate = []


rail_height = ['36"','42"','34"-38"','custom']
post_type = ['2-1/2 x 4" Base shoe','2-3/8" x 2-3/8" square aluminum posts', '2-3/8" x 2-3/8" Series 500 aluminum slotted Post','1.5" Sch 40','1.5" x 1.5" Square','1" x 2" Rectangular','custom' ]

mounting_detail = ["fascia mounted to front of deck framing using PRO's Fascia brackets", 'fascia mounted to front of deck framing using steel angle iron (angle iron installed by others)',
                    'mounted directly to deck framing using engineered lags','mounted to top of deck surface using rubber gasket and 5x5 baseplate','fascia mounted to welded knife plates (knife plates by others)',
                   'fascia mounted to angle aluminum brakets attached to halfens','mounted to stringer using 3" x 8" slotted angle base plate','1 line', '2 line','3 line','custom']

top_rail = ['Top rail profile 200', 'Top rail profile 375','Top rail profile 400','CL Laurence 1" x 1-5/16" SS  Top rail Profile','No top rail','core mounted grabrail','baseplate mounted grabrail','wall mounted grabrail','custom']

bottom_rail = ['bottom rail profile 200','with CR Laurence SS cladding','bottom rail profile 100','bottom rail profile 500','Without a Bottom rail','to be mounted directly to walls at stairs','to be mounted directly to aluminum posts at stairs',
               'to be mounted directly into core hole pockets','to be mounted directly to surface','custom']

infill = ['5/8" x 5/8" picket infill','1/4 laminated Tempered glass infill','3/8 laminated Tempered glass infill','1/2 laminated Tempered glass infill','9/16 laminated Tempered glass infill','1/8" SS Cable Infill','3/16" SS Cable infill','quikset grout',
          '4"x4" aluminum baseplate','decorative grabrail brackets','custom']

spacing = ['',"2'","3'","4'","5'","6'",'custom']

rail_type = ['Picket Guardrail','Glass Guardrail','Cable Guardrail','Grab rail','custom']

def record_custom(_input):
    cust_file = open('custom_logs.txt','a')
    cust_file.write('_input')
    cust_file.close()


def return_description(set_height,set_post,set_mount,set_top,set_bottom,set_infill,set_spacing,set_type):
    rail_description = []
    if rail_height[set_height] == 'custom':
        custom_height = input('Height: ')
        record_custom(custom_height)
        rail_description.append(custom_height)
    else:
        rail_description.append(rail_height[set_height])

    if post_type[set_post] == 'custom':
        custom_post = input('post type: ')
        record_custom(custom_post)
        rail_description.append(custom_post)
    else:
        rail_description.append(post_type[set_post])

    if mounting_detail[set_mount] == 'custom':
        custom_mount = input('mount/ GR line: ')
        record_custom(custom_mount)
        rail_description.append(custom_mount)
    else:
        rail_description.append(mounting_detail[set_mount])

    if top_rail[set_top] == 'custom':
        custom_top = input('TR / GR mount: ')
        record_custom(custom_top)
        rail_description.append(custom_top)
    else:
        rail_description.append(top_rail[set_top])

    if bottom_rail[set_bottom] == 'custom':
        custom_bottom = input('BR/ GR mount: ')
        record_custom(custom_bottom)
        rail_description.append(custom_bottom)
    else:
        rail_description.append(bottom_rail[set_bottom])

    if infill[set_infill] == 'custom':
        custom_infill = input('infill / GR mount spec: ')
        record_custom(custom_infill)
        rail_description.append(custom_infill)
    else:
        rail_description.append(infill[set_infill])

    if spacing[set_spacing] == 'custom':
        custom_spacing = input('spacing: ')
        record_custom(custom_spacing)
        rail_description.append(custom_spacing)
    else:
        rail_description.append(spacing[set_spacing])

    if rail_type[set_type] == 'custom':
        custom_type = input('rail type: ')
        record_custom(custom_type)
        rail_description.append(custom_type)
    else:
        rail_description.append(rail_type[set_type])

    return rail_description




class Excel_to_py:

    def __init__(self,file,workbook,note_sheet):
        self.file = file
        self.workbook = workbook#load_workbook(filename=file,data_only=True)
        note_sheet = workbook[note_sheet]
        self.note_sheet = note_sheet#workbook['Write up']

    def return_variables(self):
        #workbook = load_workbook(filename=file,data_only=True)
        #note_sheet = workbook['Write up']
        for row in self.note_sheet.iter_rows(min_row=1,min_col=2,max_col=2,values_only=True):
            res = ''.join(map(str,row))
            estimate.append(res)
        return estimate[0],estimate[1],estimate[2],estimate[3],estimate[4],estimate[5]

    def return_lf(self):
        section_lf= []
        for row in self.note_sheet.iter_rows(min_row=2,max_row=6,min_col=4,max_col=4,values_only=True):
            res = ''.join(map(str,row))
            if res != "None":
                if res != "0":
                    res = round(float(res),0)
                    section_lf.append(int(res))
                else:
                    section_lf.append('NA')

        return section_lf

    def return_lfprice(self):
        section_lfprice= []
        for row in self.note_sheet.iter_rows(min_row=2,max_row=6,min_col=5,max_col=5,values_only=True):
            res = ''.join(map(str,row))
            if res != 'None':
                if res != '0':
                    res = round(float(res),0)
                    section_lfprice.append(int(res))
                else:
                    section_lfprice.append('NA')

        return section_lfprice

    def return_section_details(self,num=0):
        sections = self.return_lf()
        section = []
        for row in self.note_sheet.iter_rows(min_row=2,min_col=(7+num),max_col=(7+num),values_only=True):
                res = ''.join(map(str,row))
                if res != 'None':
                    section.append(int(res))
        return section

    def return_area_name(self,num=0):
        area_names = ''
        for row in self.note_sheet.iter_rows(min_row=(2+num),max_row=(2+num),min_col=3,max_col=3,values_only=True):
            print(row)
            res = ''.join(map(str,row))
            if res != 'None':
                area_name = str(res)
                return area_name
            else:
                return 'None'

    def return_rep(self):
        for row in self.note_sheet.iter_rows(min_row=13,max_row=13,min_col=2,max_col=2,values_only=True):
            rep = ''.join(map(str,row))
            if rep.lower() == 'jag':
                return 'jag'
            elif rep.lower() == 'dave':
                return 'dave'
            else:
                return 'jag'


def write_proposal(customer_name,customer_company,contact_info,company_address,job_address,job_name,_bid_lf,_bid_lfprice):
    print('start making proposal...')
    etp = Excel_to_py(file,load_workbook(filename=file,data_only=True),'Write up')

    #load estimate number
    est_num_file = open('estimate_number.txt','r+')
    estimate_number = int(est_num_file.read())
    est_num_file.close()

    # Document Variables
    today = date.today()
    d1 = today.strftime("%m/%d/%Y")
    subtotal = 0
    total_info = [d1,customer_name,contact_info,customer_company,company_address,job_name,job_address]

    #Document Setup
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    #margin setup
    paragraph_format = style.paragraph_format
    paragraph_format.space_before = Pt(3.4)
    paragraph_format.space_after = Pt(0)
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(.5)
        section.bottom_margin = Inches(.4)
        section.left_margin = Inches(.8)
        section.right_margin = Inches(.8)

    #header setup
    header = document.sections[0].header
    h1 = header.paragraphs[0].add_run("\tPrecision Rail of Oregon, LLC\n").font.size = Pt(36)
    h2 = header.paragraphs[0].add_run("10735 SE Foster RD").font.size = Pt(14)
    header.paragraphs[0].add_run('\t').add_picture('Alumarail_logo.png', width = Inches(1.8))
    h3 = header.paragraphs[0].add_run("\tPortland, Oregon 97266").font.size = Pt(14)
    h1.italic = True
    h2.italic = True
    h3.italic = True
    h1.bold = True
    h2.bold = True
    h3.bold = True
    header.paragraphs[0].style.font.color.rgb = RGBColor(0, 96, 43)

    #footer setup
    footer = document.sections[0].footer
    footer.paragraphs[0].add_run('Phone: 503-512-5353\tFax: 503-668-8968\twww.Precisionrail.com').font.size = Pt(14)
    footer.paragraphs[0].style.font.color.rgb = RGBColor(0, 96, 43)

    # start Document
    est =document.add_paragraph()
    est.add_run('\t'*10 + 'Estimate #: {}-C'.format(estimate_number)).bold=True
    est_num = open('estimate_number.txt','w+')
    est_num.write(str(estimate_number + 1))
    est_num.close()

    #head_info = document.add_paragraph()
    for item in total_info:
        document.add_paragraph(str(item))
    document.add_paragraph('')
    document.add_paragraph('Dear ' + customer_name + ',\n')
    p1 = document.add_paragraph()
    p1.style = document.styles['Normal']

    p1.add_run('Precision Rail of Oregon is pleased to provide the following proposal for: ')
    p1.add_run(job_name + ', BUDGET Rev-0 \n\n').bold = True
    p1.add_run('Items furnished by Precision Rail of Oregon: Submittal drawings, engineering, materials, and installation.\n\n').bold = True
    p1.add_run('Submittals:').bold = True
    p1.add_run(' Pricing includes 1 submittal based off plans and 1 revision once corrections are received from GC. Any Additional revisions to be billed at 145.00 per hour plus materials and handling. \n\n')

    #sections
    for num in range(len(_bid_lf)):
        if _bid_lf[num] != 'NA':
            bid_area =etp.return_area_name(num)
            section = etp.return_section_details(num)
            b1 = return_description(section[0],section[1],section[2],section[3],section[4],section[5],section[6],section[7])
            p1.add_run('Bid Item - {} Tall {} ({})\n'.format(b1[0],b1[7],bid_area)).bold = True
            p1.add_run(" {} {}. {}, {} with {}. Posts spacing to be evenly spaced and not exceed {} as per engineering and customer request. Support blocking by others. Standard color (Black, Bronze, White). ".format(b1[1],b1[2],b1[3],b1[4],b1[5],b1[6]))
            if b1[7] == 'Grab rail':
                p1.add_run('Handrails are all ADA Compliant.')
            p1.add_run('\n\n')
            p1.add_run('\t'*6 + 'Total {} LF @ ${}.00 per LF = Sub Total ${}.00*\n\n'.format(str(_bid_lf[num]),str(_bid_lfprice[num]),str(_bid_lf[num]*_bid_lfprice[num]))).bold = True
            subtotal += _bid_lf[num] * _bid_lfprice[num]

    p1.add_run('\n'*3)
    p1.add_run('\t'*10 + '     Total = {}.00*\n\n\n'.format(str(subtotal))).bold = True


    p1.add_run('\t\t*This price quote is valid for 3 months from the date of this document*\n\n').italic = True

    p1.add_run('Assumptions\n').bold = True
    p1.add_run('The following assumptions were made in support of this estimate:')

    p1.add_run(
                """
        1.	Electrical utilities available on site.
        2.	Sanitation facilities will be provided and available on site.
        3.	Core holes, provided by others, will be cleaned out and ready for post installation.
        4.	Fall restraint anchor points will be available and cleaned out ready for use.
        5.	Paint / PPG Duracron with a 5 year warranty.
                """)
    p1.add_run('\n')
    p1.add_run('Items EXCLUDED by Precision Rail of Oregon unless noted above:')

    p1.add_run(
        """
        1.	Deferred permits or any items not specifically included is considered furnished by others.
        2.	Taxes such as sales, local municipality, gross receipts tax and/or local business licenses.
        3.	Union, prevailing wage and/or workforce training installation
        4.	Insurance requirements above and beyond: $1M/$2M (occurrence/aggregate); and $3M
            umbrella.
        5.	Performance & payment bonds.
        6.	Pollution insurance requirements.
        7.	Deviations from project plans that impede the installation of our rail as planned.
        8.	Marking / locating rebar tensions wires
        9.	Coverage / Protection of any Glazing, Brick, Building materials
        10. Inspection for testing (example UT, NDT & others)
        11. Flaggers, and / or any personnel for traffic control
        12. Lifts, swing stages, cranes, or other equipment required to install are not included in this bid
                and are to be provided by the GC.

        """)
    p1.add_run('\n\n')

    p1.add_run("Submittal drawings with approval by the representative of buyer (customer) or owner shall be considered the correct measurement and method for fabrication. Delivery schedule will be based on receipt of final approved submittal drawings.\n\nThank you for the opportunity to submit a proposal.\n\nSincerely,")
    rep = etp.return_rep()

    if rep == 'jag':
        document.add_picture('JAG_signature.png', width=Inches(2))
        sign = document.add_paragraph()
        sign.style = document.styles['Normal']
        sign.add_run('Jeff Garlitz\n')
        sign.add_run('jgarlitz@precisionrail.com\n')
        sign.add_run('541-279-8182\n')

    elif rep == 'dave':
        document.add_picture('Dave_signature.png', width=Inches(2))
        sign = document.add_paragraph()
        sign.style = document.styles['Normal']
        sign.add_run('Dave Brown\n')
        sign.add_run('Dave@precisionrail.com\n')
        sign.add_run('503-793-1972\n')


    sign.add_run('\n\n\n\n')
    sign.add_run('Acceptance of Proposal Signature _______________________              Date_______________   ')
    document.save('{}_{} - rev 0.docx'.format(customer_company,job_name))
    print('proposal finished!')



def _start():
    print('proposal writer loaded...')
    os.chdir("C:/Users/Owner/Desktop/Estimating model 1.0.7.9")
    print(os.getcwd())
    etp = Excel_to_py(file,load_workbook(filename=file,data_only=True),'Write up')
    customer_name,customer_company,contact_info,company_address,job_address,job_name = etp.return_variables()
    bid_lf = etp.return_lf()
    bid_lfprice = etp.return_lfprice()

    write_proposal(customer_name,customer_company,contact_info,company_address,job_address,job_name,bid_lf,bid_lfprice)



def file_name():
    print(os.getcwd())
    _path = askopenfilename(filetypes=[('Excel Files','*.xlsm')])
    print(_path)
    set_file(_path)



def get_cwd():
    cwd = os.getcwd()
    return cwd

window = Tk()
window.title(' Auto Proposal Writer')
window.geometry('300x100')
open_button = Button(master=window,width=20,text='Open File',command=file_name)
run_button =Button(master=window,width=20,text='Make Proposal',command=_start)
open_button.pack()
run_button.pack()
window.mainloop()
