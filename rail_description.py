

rail_height = ['36"','42"','34"-38"','custom']
post_type = ['2-1/2 x 4" Base shoe','2-3/8" x 2-3/8" square aluminum posts', '2-3/8" x 2-3/8" Series 500 aluminum slotted Post','1.5" Sch 40','1.5" x 1.5" Square','1" x 2" Rectangular','custom' ]

mounting_detail = ["fascia mounted to front of deck framing using PRO's Fascia brackets", 'fascia mounted to front of deck framing using steel angle iron (angle iron installed by others)',
                    'mounted directly to deck framing using engineered lags','mounted to top of deck surface using rubber gasket and 5x5 baseplate','fascia mounted to welded knife plates (knife plates by others)',
                   'fascia mounted to angle aluminum brakets attached to halfens','1 line', '2 line','3 line','custom']

top_rail = ['Top rail profile 200', 'Top rail profile 375','Top rail profile 400','CL Laurence 1" x 1-5/16" SS  Top rail Profile','No top rail','core mounted grabrail','baseplate mounted grabrail','wall mounted grabrail','custom']

bottom_rail = [' bottom rail profile 200',' with CR Laurence SS cladding','bottom rail profile 100','bottom rail profile 500','to be mounted directly to walls at stairs','to be mounted directly to aluminum posts at stairs',
               'to be mounted directly into core hole pockets','to be mounted directly to surface','custom']

infill = ['5/8" x 5/8" picket infill','1/4 laminated Tempered glass infill','3/8 laminated Tempered glass infill','1/2 laminated Tempered glass infill','9/16 laminated Tempered glass infill','quikset grout',
          '4"x4" aluminum baseplate','decorative grabrail brackets','custom']

spacing = ['',"2'","3'","4'","5'","6'",'custom']

rail_type = ['Picket Guardrail','Glass Guardrail','Cable Guardrail','Grab rail','custom']

def get_description():
    for num in range(len(rail_height)):
        print(str(num)+ '. for ' + rail_height[num])
    set_height = int(input(': '))
    if rail_height[set_height] == 'custom':
        rail_height[set_height] = input(': ')
        
    for num in range(len(post_type)):
        print(str(num) + '. for ' + post_type[num])
    set_post = int(input(': '))
    if post_type[set_post] == 'custom':
        post_type[set_post] = input(': ')
        
    for num in range(len(mounting_detail)):
        print(str(num) + '. for ' + mounting_detail[num])
    set_mount = int(input(': '))
    if mounting_detail[set_mount] == 'custom':
        mounting_detail[set_mount] = input(': ')
    
    for num in range(len(top_rail)):
        print(str(num) + '. for ' + top_rail[num])
    set_top = int(input(': '))
    if top_rail[set_top] == 'custom':
        top_rail[set_top] = input(': ')
        
    for num in range(len(bottom_rail)):
        print(str(num) + '. for ' + bottom_rail[num])
    set_bottom = int(input(': '))
    if bottom_rail[set_bottom] == 'custom':
        bottom_rail[set_bottom] = input(': ')
        
    for num in range(len(infill)):
        print(str(num) + '. for ' + infill[num])
    set_infill = int(input(': '))
    if infill[set_infill] == 'custom':
        infill[set_infill] = input(': ')
        
    for num in range(len(spacing)):
        print(str(num) + '. for ' + spacing[num])
    set_spacing = int(input(': '))
    if spacing[set_spacing] == 'custom':
        spacing[set_spacing] = input(': ')
        
    for num in range(len(rail_type)):
        print(str(num) + '. for ' + rail_type[num])
    set_type = int(input(': '))
    if rail_type[set_type] == 'custom':
        rail_type[set_type] = input(': ')
        
    return set_height,set_post,set_mount,set_top,set_bottom,set_infill,set_spacing,set_type

def return_description(set_height,set_post,set_mount,set_top,set_bottom,set_infill,set_spacing,set_type):
    rail_description = []
    rail_description.append(rail_height[set_height])
    rail_description.append(post_type[set_post])
    rail_description.append(mounting_detail[set_mount])
    rail_description.append(top_rail[set_top])
    rail_description.append(bottom_rail[set_bottom])
    rail_description.append(infill[set_infill])
    rail_description.append(spacing[set_spacing])
    rail_description.append(rail_type[set_type])
    return rail_description


