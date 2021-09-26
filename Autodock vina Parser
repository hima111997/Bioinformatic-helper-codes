import os, xlsxwriter

# getting the folder containing the log files
dir_ = input('enter the destinations containing the log file form VINA: ')

# getting the names of the log files
logs = [f for f in os.listdir(dir_) if f.endswith('.log')]

# a dict with name of the file as key, and a value as a list with the docking affinities
drug_affinity = {}

errors = [] # if some files produced an error they will be added here

# for each log file, open it, read it, find all conformations, get the docking affinity,
# then add it in drug_affinity dict
for log in logs:
    with open('{}/{}'.format(dir_, log)) as f:
        data = f.read()
    idx_start = data.find('   1    ')
    idx_end = len(data)
    affinities = []
    res = data[idx_start:idx_end]
    #print(log)
    #print(res)
    modes = res.splitlines()    
    #print(modes)
    for mode in modes:
        try:
            affinity = mode.split()[1]
        except:
            print('error with {}'.format(log))
            errors.append(log)
            continue
        affinities.append(affinity)
    drug_affinity[log.split('.')[0]] = affinities
    #print(drug_affinity)
    #break

# convert the drug_affinity into a list and reorder it using the first conformation affinity
list_sorted = list(drug_affinity.items())
list_sorted.sort(key = lambda x: float(x[1][0]))


# write the data in an excel file
xlsx_name = input('\n\nEnter the name of the excel file: ')
workbook = xlsxwriter.Workbook(dir_+'\\{}.xlsx'.format(xlsx_name))
worksheet = workbook.add_worksheet()
for idx, (mode, affs) in enumerate(list_sorted):
    worksheet.write(idx, 0, int(mode))
    for idx_affs, aff in enumerate(affs):
        worksheet.write(idx, idx_affs+1, float(aff))
    
workbook.close()
print('\n\nfiles produced error are: ',errors)
print('\n\nparsing data finished')
