'''function for calculating project baselines'''


from bcompiler.utils import project_data_from_master
from openpyxl import Workbook
import datetime
from openpyxl.chart import ScatterChart, Reference, Series

''' function to returns a dictionary structured in the following way project name[('latest quarter info', 'latest bc'), 
('last quarter info', 'last bc'), ('last baseline quarter info', 'last baseline bc'), ('oldest quarter info', 
'oldest bc')] depending on the amount information available in the data. Only the first three key values are returned, 
to ensure consistency (which is helpful later).'''
def baselining(project_name_list, master_list):

    output_dict = {}

    for name in project_name_list:
        all_list = []      # format [('quarter info': 'bc')] across all masters including project
        bl_list = []        # format ['bc', 'bc'] across all masters. bl_list_2 removes duplicates
        ref_list = []       # format as for all list but only contains the three tuples of interest
        for master in master_list:
            try:
                bc_stage = master[name]['BICC approval point']
                quarter = master[name]['Reporting period (GMPP - Snapshot Date)']
                tuple = (quarter, bc_stage)
                all_list.append(tuple)
            except KeyError:
                pass

        for i in range(0, len(all_list)):
            bl_list.append(all_list[i][1])


        '''below lines of text from stackoverflow. Question, remove duplicates in python list while 
        preserving order'''
        seen = set()
        seen_add = seen.add
        bl_list_final = [x for x in bl_list if not (x in seen or seen_add(x))]

        ref_list.insert(0, all_list[0])     # puts the latest info into the list first

        try:
            ref_list.insert(1, all_list[1])    # puts that last info into the list
        except IndexError:
            ref_list.insert(1, all_list[0])

        if len(bl_list_final) == 1:                     # puts oldest info into list (if no baseline)
            ref_list.insert(2, all_list[-1])
        else:
            for i in range(0, len(all_list)):      # puts in baseline
                if all_list[i][1] == bl_list_final[1]:
                    ref_list.insert(2, all_list[i])

        '''there is a hack here i.e. returning only first three in ref_list. There's a bug which I don't fully 
        understand, but this solution is hopefully good enough for now'''
        output_dict[name] = ref_list
        #[0:3]

    return output_dict

'''master files are loaded here'''
zero_2 = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2018.xlsx')
zero = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_2_2018.xlsx')
one = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_1_2018.xlsx')
two = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_4_2017.xlsx')
three = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2017.xlsx')
four = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_2_2017.xlsx')
five = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_1_2017.xlsx')
six = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_4_2016.xlsx')
last = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2016.xlsx')

'''master files put into a list'''
master_list_all = [zero_2, zero, one, two, three, four, five, six, last]
#master_list_two = [zero, last]

project_names = zero.keys()
project_names_one = ['South Eastern Rail Franchise Competition']

milestone_keys = ['Project MM18', 'Approval MM1', 'Approval MM3', 'Approval MM10', 'Project MM19', 'Project MM20', 'Project MM21']
milestone_dates = ['Project MM18 Forecast - Actual', 'Approval MM1 Forecast / Actual', 'Approval MM3 Forecast / Actual', 'Approval MM10 Forecast / Actual',
                   'Project MM19 Forecast - Actual', 'Project MM20 Forecast - Actual', 'Project MM21 Forecast - Actual']

list_of_bc_stages = baselining(project_names, master_list_all)

#milestone_data = milestone_extraction(project_names, master_list_all, milestone_keys, milestone_dates)

#td_data = cal_td(milestone_data)

