#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar  5 11:40:35 2025

@author: ijeong-yeon
"""
import pandas as pd
import bodhi_monitoring as bd

# Please assign the directory of dataset (Always place the dataset in the data folder)
data_file_directory = "data/raw_updated.xlsx"
df = pd.read_excel(data_file_directory)

# If the project has conducted any pilot data, please specify here: ['2025-03-10', '2025-03-11', etc]
pilot_test_dates = []

# Please specify the column name of enumerators. If the survey has more than 1, please add all of them
# For example: ['Enumerator Name (Kigali)', 'Enumerator Name (Southern Province)', etc]
enumerator_names = ['Enumerator Name']

# Please specify the column name of location
# For example: 'A2-2. Province'
location = None

# Please specify the column name of gender
# For example: 'A4. Gender'
gender = None

# Please specify the column name of age
# For example: 'A5. What is your age (in years)?'
age = None

# Please specify the column names of WG-SS scale
disability = ['Col1','Col2','Col3','Col4','Col5','Col6']

# If the team needs to collect different type of respondent, please specify the column for the respondent type
# For example: "0. Which stakeholder is being interviewed?"
respondent_type = None


# Please update 'project_name' to Bodhi's internal project name
project_name = bd.Bodhi_monitoring('project_name', df)
project_name.setting(pilot_test_dates, enumerator_names, location, gender, age, disability,respondent_type)
project_name.run()