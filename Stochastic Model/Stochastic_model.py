import openpyxl
import pandas as pd
from datetime import datetime,timedelta
import gurobipy as grb
from gurobipy import Model, GRB, quicksum
import numpy as np
import ProbEstimates_r as pr_requests
import ProbEstimates_c as pr_cancellations
import GetRev as rev
import time

hotelData_requests = pd.read_excel('hotelData_requests.xls')
hotelData_cancellations = pd.read_excel('hotelData_cancellations.xls')

start_time = time.time()

m = grb.Model("Hotel_Booking")

# Increase verbosity of the output
m.setParam('OutputFlag', 1)             # Ensure output is turned on
m.setParam('DisplayInterval', 1)        # Display log every one second

# Log to console and to a file
m.setParam('LogToConsole', 1)           # Display log on console
m.setParam('LogFile', 'gurobi_log.txt') # Write log to a file

# Tune heuristic visibility (optional)
m.setParam('Heuristics', 1.0)  

m.setParam('MIPFocus', 1)    # Focus on finding feasible solutions quickly
m.setParam('Cuts', 3)        # More aggressive cut generation

room_combinations = {
    'CR': [(1, 0), (1, 1), (2, 0)],
    'JCR': [(1, 0), (1, 1), (2, 0)],
    'CCR': [(1, 0), (1, 1), (2, 0)],
    'DCS': [(1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (2, 2), (3, 0)],
    'SIS': [(1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (2, 2), (3, 0)],
    'SS': [(1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (2, 2), (3, 0)],
    'WS': [(1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (2, 2), (3, 0)],
    'LS': [(1, 0), (1, 1), (1, 2), (1, 3), (2, 0), (2, 1), (2, 2), (3, 0)]
}

room_classes = [1, 2, 3, 4, 5, 6, 7, 8]
U_i = {1: 5, 2: 2, 3: 3, 4: 2, 5: 1, 6: 2, 7: 1, 8: 1}

curlyC = {"CR", "JCR", "CCR", "DCS", "SIS", "SS", "WS", "LS"}

room_type_to_int = {
    "CR" : 1,
    "JCR": 2,
    "CCR": 3,
    "DCS": 4,
    "SIS": 5,
    "SS" : 6,
    "WS" : 7,
    "LS" : 8,}

t_instances = 4

#Dates of reservation interval
start_reservation_date_str = "2023-12-20"
end_reservation_date_str = "2023-12-21"


#Dates of service interval
start_service_date_str = "2023-12-22"
end_service_date_str = "2023-12-23"


#creating a reservation object including all the parameters
class Reservation:
    def __init__(self, rrt, ad, length, na, nc):
        self.rrt = rrt
        self.ad = ad
        self.length = length
        self.na = na
        self.nc = nc

    def __hash__(self):
        # Define a custom hash function based on attributes
        return hash((self.rrt, self.ad, self.length, self.na, self.nc))

    def __eq__(self, other):
        if not isinstance(other, Reservation):
            return False
        return self.rrt == other.rrt and self.ad == other.ad and self.length == other.length and self.na == other.na and self.nc == other.nc
    
    def __str__(self):
        return f'rrt: {self.rrt}, ad: {self.ad}, length: {self.length}, na: {self.na}, nc: {self.nc}'
    
    #in order to get the arrival date from the object 
    def getDateArrival(self):
        date_object = datetime.strptime(self.ad, '%Y/%m/%d')
        day_of_year = date_object.timetuple().tm_yday
        return day_of_year
    
    #in order to get the room type from the object 
    def getRoomType(self):
        switcher = {
            "CR" : 1,
            "JCR": 2,
            "CCR": 3,
            "DCS": 4,
            "SIS" : 5,
            "SS": 6,
            "WS" : 7,
            "LS" : 8,
        }
        return switcher.get(self.rrt, "nothing")
    
curlyK = []

for room_class in curlyC:
    if room_class in room_combinations:
        for ad in pd.date_range(start=start_service_date_str, end=end_service_date_str).strftime('%Y/%m/%d'):
            for length in range(1, 4):  # Assuming length can be between 1 and 14
                for combination in room_combinations[room_class]:
                    na, nc = combination  # Iterate over valid combinations
                    newReservation = Reservation(room_class, ad, length, na, nc)
                    curlyK.append(newReservation)

#Finding curlyT
def date_to_day_of_year(date_str):
    #Convert a date in string format to its day number in the year.
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.timetuple().tm_yday

def date_to_day_of_year2(date_str):
    #Convert a date in string format to its day number in the year.
    date_obj = datetime.strptime(date_str, "%Y/%m/%d")
    return date_obj.timetuple().tm_yday

#Reservation Interval

def reservation_date_to_time_periods(start_reservation_date_str, end_reservation_date_str):
  
    #Convert a reservation date range to the corresponding time periods.Returns a list of time periods.
   
    start_day = date_to_day_of_year(start_reservation_date_str)
    end_day = date_to_day_of_year(end_reservation_date_str)
    
    time_periods = []
    
    for day in range(start_day, end_day + 1):  # +1 to include the end day
        for period in range(1, t_instances + 1):
            t = (day - 1) * t_instances + period  # Calculate time period number for this day and period
            time_periods.append(t)
    
    return time_periods


curlyT_actual = reservation_date_to_time_periods(start_reservation_date_str, end_reservation_date_str)

def normalize_list(curlyT_actual):
    if not curlyT_actual:
        return []
    min_val = min(curlyT_actual)
    return [x - min_val + 1 for x in curlyT_actual]

curlyT = normalize_list(curlyT_actual)

#Finding Tau
def get_last_time_period_of_reservation(start_reservation_date_str, end_reservation_date_str):
    time_periods = reservation_date_to_time_periods(start_reservation_date_str, end_reservation_date_str)
    curlyT = normalize_list(time_periods)
    return curlyT[-1] if curlyT else None

tau = get_last_time_period_of_reservation(start_reservation_date_str, end_reservation_date_str)

# Finding curlyL
def service_date_to_days(start_date_str, end_date_str):
    
    # Convert a service date range to the corresponding days in the year. Returns a list of days.
    start_day = date_to_day_of_year(start_date_str)
    end_day = date_to_day_of_year(end_date_str)
    
    return list(range(start_day, end_day + 1))  # +1 to include the end day

#Service interval

service_days = service_date_to_days(start_service_date_str, end_service_date_str)

def normalize_days(service_days):
    if not service_days:
        return []
    min_day = min(service_days)
    max_day = max(service_days)
    return [day - min_day + 1 for day in range(min_day, max_day + 1)]

curlyL = normalize_days(service_days)

#since the GetProbability function reads through the full name.
#hotelData.xls room type cells contains the full name. 
def getRoom(C):
    switcher = {
        "CR": "Comfort Room",
        "JCR": "Junior Comfort Room",
        "CCR": "Corner Comfort Room",
        "DCS": "Deluxe Comfort Suite",
        "SS": "Superior Suite",
        "SIS": "Signature Suite",
        "WS": "Wellness Suite",
        "LS": "Luxury Suite",
    }
    return switcher.get(C, "nothing")

#dates of the seasons
seasons = {
    "LS": [
        ("2021-11-01", "2021-12-20"),
        ("2022-01-07", "2022-04-10"),
        ("2022-11-01", "2022-12-20"),
        ("2023-01-07", "2023-04-10"),
        ("2023-11-01", "2023-12-20"),
        ("2024-01-07", "2024-04-10"),
        ("2024-11-01", "2024-12-20"),
        ("2025-01-07", "2025-04-10"),
        ("2025-11-01", "2025-12-20"),
    ],
    "SS": [
        ("2021-10-01", "2021-10-31"),
        ("2021-12-21", "2022-01-06"),
        ("2022-04-11", "2022-05-31"),
        ("2022-10-01", "2022-10-31"),
        ("2022-12-21", "2023-01-06"),
        ("2023-04-11", "2023-05-31"),
        ("2023-10-01", "2023-10-31"),
        ("2023-12-21", "2024-01-06"),
        ("2024-04-11", "2024-05-31"),
        ("2024-10-01", "2024-10-31"),
        ("2024-12-21", "2025-01-06"),
        ("2025-04-11", "2025-05-31"),
        ("2025-10-01", "2025-10-31"),
        ("2025-12-21", "2026-01-06"),
    ],
    "HS": [
        ("2021-09-07", "2021-09-30"),
        ("2022-06-01", "2022-06-30"),
        ("2022-09-01", "2022-09-30"),
        ("2023-06-01", "2023-06-30"),
        ("2023-09-01", "2023-09-30"),
        ("2024-06-01", "2024-06-30"),
        ("2024-09-01", "2024-09-30"),
        ("2025-06-01", "2025-06-30"),
        ("2025-09-01", "2025-09-30"),
    ],
    "VHS": [
        ("2022-07-01", "2022-07-31"),
        ("2022-08-01", "2022-08-31"),
        ("2023-07-01", "2023-07-31"),
        ("2023-08-01", "2023-08-31"),
        ("2024-07-01", "2024-07-31"),
        ("2024-08-01", "2024-08-31"),
        ("2025-07-01", "2025-07-31"),
        ("2025-08-01", "2025-08-31"),
    ]
}


#Give the arrival-date season
def get_season(arrival_date):
    #To check in what season is arrival_date
    arrival_date = datetime.strptime(arrival_date, "%Y/%m/%d")
    for season, date_ranges in seasons.items():
        for start, end in date_ranges:
            start_date = datetime.strptime(start, "%Y-%m-%d")
            end_date = datetime.strptime(end, "%Y-%m-%d")
            if start_date <= arrival_date <= end_date:
                return season
    return "Invalid Date"
    

def get_suite(roomtype):
    #If Not Suite Return False
    if(roomtype == "DCS" or roomtype == "SIS" or roomtype == "SS"  or roomtype == "WS" or roomtype == "LS"):
        return True
    else:
        return False


def is_weekend(ad):
    date_object = datetime.strptime(ad, '%Y/%m/%d')
    # weekday() returns 0 for Monday, 4 for Friday, and 6 for Sunday
    if date_object.weekday() in [4, 5, 6]:
        return True
    else:
        return False


def get_month_from_time_index_t(t, t_instances):
    
    time_period = curlyT_actual[t-1]
    # Number of days in each month
    days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    # Calculate boundaries based on t_instances
    boundaries = {}
    start = 1
    for month, days in enumerate(days_in_month, 1):
        end = start + days * t_instances - 1
        boundaries[month] = (start, end)
        start = end + 1

    # Define month names
    months = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December"
    }
    

    # Iterate over the boundaries to determine the month
    for month, (start, end) in boundaries.items():
        if start <= time_period <= end:
            return months[month]

    return "Invalid t value"


def days_in_given_month(month_name):
    month_to_days = {
        "January": 31,
        "February": 28,  # This doesn't account for leap years
        "March": 31,
        "April": 30,
        "May": 31,
        "June": 30,
        "July": 31,
        "August": 31,
        "September": 30,
        "October": 31,
        "November": 30,
        "December": 31
    }
    
    return month_to_days.get(month_name, 0)


def getProb_requests(k, t):
    # Getting room type from reservation list
    room_type = getRoom(k.rrt)
    
    # Getting month from time_index set
    month_name = get_month_from_time_index_t(t, t_instances)
    
    # Getting number of days in the month
    days_in_month = days_in_given_month(month_name)
    
    # Getting season type from arrival date
    season = get_season(k.ad)
    
    # Getting inSuite from reservation list
    IsSuite = get_suite(k.rrt)
    
    IsWeekend = is_weekend(k.ad)
    
    # print(k.ad)
    probability = pr_requests.calculate_prob_requests(hotelData_requests, room_type, month_name, season, k.length, IsWeekend, IsSuite, k.na, k.nc)
    
    # Divide the probability by the number of time periods in the month
    total_time_periods = days_in_month * t_instances
    if total_time_periods != 0:  # Avoid division by zero
        return probability / total_time_periods
    
    return 0  # In case of an error, return 0


def getProb_cancellations(k, t):
    # Getting room type from reservation list
    room_type = getRoom(k.rrt)
    
    # Getting month from time_index set
    month_name = get_month_from_time_index_t(t, t_instances)
    
    # Getting number of days in the month
    days_in_month = days_in_given_month(month_name)
    
    # Getting season type from arrival date
    season = get_season(k.ad)
    
    # Getting inSuite from reservation list
    IsSuite = get_suite(k.rrt)
    
    IsWeekend = is_weekend(k.ad)
    
    # print(k.ad)
    probability = pr_cancellations.calculate_prob_cancellations(hotelData_cancellations, room_type, month_name, season, k.length, IsWeekend, IsSuite, k.na, k.nc)
    
    # Divide the probability by the number of time periods in the month
    total_time_periods = days_in_month * t_instances
    if total_time_periods != 0:  # Avoid division by zero
        return probability / total_time_periods
    
    return 0  # In case of an error, return 0


def getAik(c, i, k):
    arrival_day = date_to_day_of_year2(k.ad)
    length_of_stay = k.length
        
    if k.rrt == c and arrival_day <= service_days[i-1] <= arrival_day + length_of_stay - 1:
        return True
    else:
        return False


#Decision variables

u = {}               
for k in curlyK:    
    for t in range(1,tau+1):
        u[k, t] = m.addVar(vtype=GRB.BINARY, name=f'u_{k}_{t}')
                            
r = {}
for c in curlyC:  # Loop over room classes
    for i in curlyL:  # Loop over days
        for t in range(1, tau+2):  # Loop over time periods
            # Extract the upper bound for the room type from U_i
            upper_bound = U_i[room_type_to_int[c]]
            r[c, i, t] = m.addVar(vtype=grb.GRB.INTEGER, lb=0, ub=upper_bound, name=f'r_{c}_{i}_{t}')
        

# Constraints

# Initialization for t=0

# Constraint 1 - Initial Room Availability
for c in curlyC:  # Loop over room classes
    for i in curlyL:  # Loop over days
        m.addConstr(r[c, i, 1] == U_i[room_type_to_int[c]])


  # Constraint 2 - Reservation Acceptance Constraint
for c in curlyC:  # Loop over room classes
    for i in curlyL:  # Loop over days
        for t in range(2, tau + 2):  # Start from the second time period
            m.addConstr(
                r[c, i, t] == r[c, i, t-1] - quicksum(getAik(c, i, k) * u[k, t-1] for k in curlyK))
            
# for k in curlyK:
#     for t in curlyT:
#         m.addConstr(getProb_requests(k, t) > 0)

            
    
m.setObjective(quicksum(rev.GetRev(k)* getProb_requests(k, t) * (1 - getProb_cancellations(k, t)) * u[k, t] for k in curlyK for t in curlyT), GRB.MAXIMIZE)
m.optimize()


method_used = m.Params.Method
print("Algorithm used:", method_used)

if m.status == grb.GRB.Status.OPTIMAL:
    total_revenue = 0  # Initialize the total revenue to 0
    for k in curlyK:
        for t in curlyT:
            if u[k, t].x > 0:  # Check if the reservation is accepted
                probability_kt = getProb_requests(k, t)
                
                if probability_kt > 0:  # Check if the probability is strictly positive
                    revenue_for_k = rev.GetRev(k)
                    total_revenue += revenue_for_k  # Accumulate the revenue
                    month_name = get_month_from_time_index_t(t, t_instances)

                    # Print details of the accepted reservation
                    print(f'Reservation product rrt: {k.rrt}, ad: {k.ad}, Reservation Month: {month_name}, length: {k.length}, na: {k.na}, nc: {k.nc} ' 
                          f'is accepted in time instance {t} with revenue {revenue_for_k:.2f} and probability {probability_kt:.6f}.')
    
    print(f'Total revenue: {total_revenue:.2f}')  # Print the total revenue
else:
    print('No solution found.')

end_time = time.time()
duration = end_time - start_time

minutes, seconds = divmod(duration, 60)
print(f"\nThe Gurobi optimization took {int(minutes)} minutes and {seconds:.2f} seconds to complete.")

# if m.status == grb.GRB.Status.OPTIMAL:
#     total_revenue = 0  # Initialize the total revenue to 0
#     for k in curlyK:
#         for t in curlyT:
#             if u[k, t].x > 0:
#                 revenue_for_k = rev.GetRev(k)
#                 total_revenue += revenue_for_k 
#                 probability_kt = getProb_requests(k, t)# Accumulate the revenue
#                 month_name = get_month_from_time_index_t(t, t_instances)

                
#                 # Modified print statement to match desired format and include probability
#                 print(f'Reservation product rrt: {k.rrt}, ad: {k.ad}, Month name: {month_name},length: {k.length}, na: {k.na}, nc: {k.nc} ' 
#                       f'is accepted in time instance {t} with revenue {revenue_for_k:.2f} and probability {probability_kt:.6f}.')
                
#     print(f'Total revenue: {total_revenue:.2f}')  # Print the total revenue
# else:
#     print('No solution found.')


                
                

   # if m.status == grb.GRB.Status.OPTIMAL:
   #     total_revenue = 0  # Initialize the total revenue to 0
   #     for k in curlyK:
   #         for t in curlyT:
   #             if u[k, t].x > 0:
   #                 revenue_for_k = rev.GetRev(k)
   #                 total_revenue += revenue_for_k  # Accumulate the revenue
   #                 print(f'Reservation product {k} is accepted in time period {t} with revenue {revenue_for_k}.')
   #     print(f'Total revenue: {total_revenue}')  # Print the total revenue
   # else:
   #     print('No solution found.')


   # # Check the algorithm used
   # method_used = m.Params.Method
   # print("Algorithm used:", method_used)             
                
                
                
                
                
                
                
       # #Constraint 3 - Reservation Acceptance Constraint
       # for i in curlyL:   
       #     for t in range(2,tau+2):
       #         m.addConstr(quicksum(getAik(i, k) * u[k, t-1] for k in curlyK) <= r[i,t])         
                
                
                
                
                
