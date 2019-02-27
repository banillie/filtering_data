'''

planning on using this as a place to have helper functions for all filter functions. It the moment it contains a early
dev function that places project names in a specific order, based on meta data of interest. However, wondering it might
be better to simply place this information into ws so can be filtered there.

'''


from openpyxl import Workbook
from bcompiler.utils import project_data_from_master
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
import random


def place_in_order(data, category):
    initial_category_list = []

    for project_name in data:
        initial_category_list.append(data[project_name][category])

    final_cat_list = list(set(initial_category_list))

    name_order = []
    for cat in final_cat_list:
        for project_name in data:
            if data[project_name][category] is cat:
                name_order.append(project_name)

    return name_order


q3_1819 = project_data_from_master('C:\\Users\\Standalone\\Will\\masters folder\\master_3_2018.xlsx')


a = place_in_order(q3_1819, 'DfT Group')

'''def filter_group(dictionary, group_of_interest):
    project_list = []
    for project in dictionary:
        if dictionary[project]['DfT Group'] == group_of_interest:
            project_list.append(project)

    return project_list'''