def filter_group(dictionary, group_of_interest):
    project_list = []
    for project in dictionary:
        if dictionary[project]['DfT Group'] == group_of_interest:
            project_list.append(project)

    return project_list