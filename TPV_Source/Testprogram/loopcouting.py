import random

intervalls = ['Daglig', 'Daglig', 'Daglig', 'Ukentlig', 'Ukentlig', 'Månedlig', 'Månedlig', 'Halvår', 'Halvår']

dag1 = [1, 1, 1, 1, 1, 0, 0, 0, 0]
dag2 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag3 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag4 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
dag5 = [1, 1, 1, 0, 0, 0, 0, 0, 0]
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
# no signal should be omitted for the first 4 list in this test

def maintainance_check(yesterday, today, maintainance, intervall):

    #instead of going by dates here, I just use a bool
    #maintainance_day = False

    for i in ukentlig.values():
        
        # How do we decide that maintainance isnt done????
        if yesterday[i] == 0 and not maintainance:
            print("Hello, you should have done some maintainance yesterday!")

        else:
            print("Everything seems to be fine")

# Made the bool a part of the function call instead of defining it in the function
# but we might be on to something here
maintainance_check(dag6, dag7, False, 0)

# one idea is to get yesterdays values from the excel document and check and see if today is a "maintainace" day
# but this in it self doesnt solve the issue, we still dont know if the maintainance was done on time.
# So I think we need either a persistent variable that is getting set when the maintainance has been done
# or we need to check backwards and calculate when the last maintainace should have been done
# and make the labels glow until it is done.
# This should also only be done if we notice that it wasnt done on time
# The last option seems to be the easiest one