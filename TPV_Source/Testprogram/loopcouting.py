import random

intervalls = ['Daglig', 'Daglig', 'Daglig', 'Ukentlig', 'Ukentlig', 'Månedlig', 'Månedlig', 'Halvår', 'Halvår']

dag1 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag2 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag3 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag4 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag5 = [1, 1, 1, 1, 1, 0, 0, 0, 0]
dag6 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag7 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag8 = [1, 1, 1, 0, 0, 0, 0, 0, 0]

key_ukentlig = 0
ukentlig = {}

key_månedlig = 0
månedlig = {}

key_halvår = 0
halvår = {}

intervalls_counter = 0


# core loop that populates the window with labels.
# we can extract info on which column each interval is in
# but are saving them in a dict the best option?
for i in intervalls:

    if i == 'Daglig':
        print('Daglig')
    elif i == 'Ukentlig':
        ukentlig['Ukentlig' + str(key_ukentlig)] = intervalls_counter
        key_ukentlig += 1
        print(ukentlig)
    elif i == 'Månedlig':
        månedlig['Månedlig' + str(key_månedlig)] = intervalls_counter
        key_månedlig += 1
        print(månedlig)
    elif i == 'Halvår':
        halvår['Halvår' + str(key_halvår)] = intervalls_counter
        key_halvår += 1
        print(halvår)
    
    intervalls_counter += 1

# Concive a way to check the 8 sample list and see if the correct maintainance is done
# if its not done, emit a signal of some sort for each passing list which dosent have the maintainace done
# after it should have been done
# no signal should be omitted for the first 4 list