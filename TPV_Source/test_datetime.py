# -*- coding: utf-8 -*-
"""
Created on Thu Aug 31 16:08:30 2017

@author: Iver
"""

from datetime import datetime, date


def test(n):
    
    date1 = datetime.today()
    year = date1.strftime('%Y')
    year = int(year)
    
    month = date1.strftime('%m')
    month = int(month)
    
    day = n
    
    check_day = date(year, month, day).isoweekday()
    
    if check_day == 7:
        day = day + 1
        
    elif check_day == 6:
        day = day - 1
        
    else:
        print('Shit yo!')
        
        
        
#        if saved_day == 20 and check_day != 6 or saved_day == 20 and check_day != 7:
#            print('Returning 20')
#            return 20
#        
#        elif saved_day == 19 and check_day != 7:
#            print('Returning 19')
#            return 19
#        
#        elif saved_day == 21 and check_day != 6:
#            print('Returning 21')
#            return 21
#        
#        else:
#            
#            if check_day == 7:
#                day = day + 1
#                return day
#                
#            elif check_day == 6:
#                day = day - 1
#                return day
#                
#            else:
#                print('Shit yo!')
    
    




