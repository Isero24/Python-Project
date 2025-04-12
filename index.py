import pandas as pd
from openpyxl import load_workbook
import random

# Function that generates a random key from a list
# Recursively runs until value is generated that's not in the excludes arg
def random_exclude(excludes):
    r = random.choice(list(workbook.keys()))
    if r in excludes:
        return random_exclude(excludes)
    return r

# Excel file is loaded
file_path = 'Test.xlsx'
workbook = pd.read_excel(file_path, sheet_name=None, header=None)

# Max number of trials that can be selected from each set
num_of_trials_per_set = 5

for x in range(1,100):

    # Generates the random trials from each set
    all_excluded_sets = []
    i = 0
    while i < len(workbook)*num_of_trials_per_set:

        # Sometimes a nonvalid randomized set is generated
        # If that happens, the process is reset and randomization is reinitialized
        if (len(set(all_excluded_sets)) == len(workbook)):
            print("Failed to generate legal randomization")
            print("resetting randomization")
            i = 0

        # intializing
        if (i == 0):
            set_max_count = {}
            for sets in workbook:
                set_max_count[sets] = 0

            exclude_previous_set = []
            exclude_max_use_sets = []
            all_excluded_sets = []
            initial_trials = workbook.copy()
            random_trials = pd.DataFrame()

        random_set = random_exclude(all_excluded_sets)
        exclude_previous_set = random_set

        # Randomely select a trial from the trial set
        # Write the trial to a new datafile
        # Remove trial from initial set
        index = random.choice(list(initial_trials[random_set].index[1:]))
        new_row = initial_trials[random_set].iloc[[index]]
        new_row.insert(0, 'TrialSetSize', random_set)
        random_trials = pd.concat([random_trials, new_row], ignore_index=True)
        initial_trials[random_set] = initial_trials[random_set].drop(index)
        initial_trials[random_set] = initial_trials[random_set].reset_index(drop=True)



        if (set_max_count[random_set] < num_of_trials_per_set): 
            set_max_count[random_set] += 1

        if (set_max_count[random_set] == num_of_trials_per_set):
            exclude_max_use_sets.append(random_set)

        all_excluded_sets = exclude_max_use_sets + [exclude_previous_set]

        i += 1
    sheet_name='Random Trials ' + str(x)
    with pd.ExcelWriter("Results.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        random_trials.to_excel(writer, sheet_name=sheet_name, index=False)

print("Finished")