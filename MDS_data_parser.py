import xlsxwriter, os


# Input the destination folder containing the data file for MDS
dir_ = input('Enter the destenation containing the data files: ')
name = dir_.split('\\')[-1]

# Create an new Excel file with the name of the folder and add a worksheet.
workbook = xlsxwriter.Workbook(dir_+'\\{}.xlsx'.format(name))
worksheet = workbook.add_worksheet()


def cleaning(data, name):
    '''Cleaning MDS data.
    Inputs:
        data: a multi line string
        name: name of the file
    Outputs:
        a clean list of numbers only
    '''
    if ' ' in data:
        data_list = data.strip().replace('\n',' ').split(' ')        
        final_data = []
        if 'bonds' in name or 'RMSF' in name:
            for idx, val in enumerate(data_list):            
                if idx%2 == 1:
                    final_data.append(val)            
        if 'gyr' in name:
            final_data = data_list[2:]            
    else:
        final_data = data.strip().split('\n')
    
    return final_data

def set_chart(chart, title, x_title, y_title=None):
    '''Modifying the chart.
    Inputs:
        chart: chart object
        title: name of the chart title
        x_title: name of the X bar
        y_title (optional): name of the Y bar
    Outputs:
        None
    '''
    chart.set_title({        
        'name': title,
        'name_font': {            
            'name': 'Calibri',
            'color': 'black',
                      },
                    })
    #print(title, x_title)
    chart.set_x_axis({
        'name': x_title,
        'name_font': {
            'name': 'Calibri',
            'color': 'black',
                      },
        'num_font': {
            'name': 'Calibri',
            'color': 'black',
                      },
                    })

#     chart.set_y_axis({
#         'name': y_title,
#         'name_font': {
#             'name': 'Calibri',
#             'color': 'black',
#                       },
#         'num_font': {
#             'name': 'Calibri',
#             'color': '#7030A0',
#                       },
#                     })


# name of the columns
COL_NAMES='ABCDEFGHIJK'

# getting the names of the data files without .dat extension
names = [f[:-4] for f in os.listdir(dir_) if f.endswith('.dat') and 'MMPBSA' not in f]

# For each file, 1) open it, 2) read it, 3) clean it, 4) convert the numbers to floats, 
# 5) add the name of the file in the clened list,
# then write the data and add the chart in the excel sheet
for idx, n in enumerate(names):
    with open(dir_+'\\'+n+'.dat') as data_file:  # (1)
        data = data_file.read()                  # (2)
        
    data = list(map(float, cleaning(data, n)))   # (3,4)
    data.insert(0, n)                            # (5)
    
    worksheet.write_column(0, idx, data)
    chart = workbook.add_chart({'type': 'line'})
    
    title = n
    if 'RMSF' in n :
        x_title = 'Amino Acid Number'
    else:        
        x_title = 'Frame Number'
    #y_title = 
    #print(n,x_title)
    set_chart(chart, title, x_title)
    
    chart.add_series({'values': '=Sheet1!${}$2:${}${}'.format(COL_NAMES[idx], COL_NAMES[idx], len(data)),})
    chart.set_legend({'position': 'none'})
    worksheet.insert_chart(20, idx*5, chart)
    

workbook.close()
print('\n\nFinished parsing data. \nyou can find the result in: {} folder'.format(dir_+'\\'+name))
