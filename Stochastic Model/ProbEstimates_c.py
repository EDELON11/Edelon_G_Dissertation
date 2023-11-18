import openpyxl
import pandas as pd
from datetime import datetime
# Read the Excel file
hotelData_cancellations = pd.read_excel('hotelData_cancellations.xls')
hotelData_requests = pd.read_excel('hotelData_requests.xls')

#date ranges for different seasons
seasons = {
    "LS": [
        ("2021-11-01", "2021-12-20"),
        ("2022-01-07", "2022-04-10"),
        ("2022-11-01", "2022-12-20"),
        ("2023-01-07", "2023-04-10")
    ],
    "SS": [
        ("2021-10-01", "2021-10-31"),
        ("2021-12-21", "2022-01-06"),
        ("2022-04-11", "2022-05-31"),
        ("2022-10-01", "2022-10-31"),
        ("2022-12-21", "2023-01-06"),
        ("2023-04-11", "2023-05-31")
    ],
    "HS": [
        ("2021-09-07", "2021-09-30"),
        ("2022-06-01", "2022-06-30"),
        ("2022-09-01", "2022-09-30"),
        ("2023-06-01", "2023-06-30"),
        ("2023-09-01", "2023-09-30"),
    ],
    "VHS": [
        ("2022-07-01", "2022-07-31"),
        ("2022-08-01", "2022-08-31"),
        ("2023-07-01", "2023-07-31"),
        ("2023-08-01", "2023-08-31")
    ]
}

def get_season(arrival_date):
    arrival_date = datetime.strptime(arrival_date, "%Y-%m-%d %H:%M:%S")
    for season, date_ranges in seasons.items():
        for start, end in date_ranges:
            start_date = datetime.strptime(start, "%Y-%m-%d")
            end_date = datetime.strptime(end, "%Y-%m-%d")
            if start_date <= arrival_date <= end_date:
                return season
    return "Invalid Date"


def number_to_month(number):
    months = {
        1: 'January',
        2: 'February',
        3: 'March',
        4: 'April',
        5: 'May',
        6: 'June',
        7: 'July',
        8: 'August',
        9: 'September',
        10: 'October',
        11: 'November',
        12: 'December'
    }
    return months.get(number, 'Invalid Month')


def getMonth(date):
    date_time_obj = datetime.strptime(str(date), '%Y-%m-%d %H:%M:%S')    
    month = date_time_obj.month
    return (number_to_month(month))
            
def calculate_prob_cancellations(hotelData_cancellations, requested_room_type,
                          reservation_date_month,arrival_date_season,
                          length_of_stay,weekend_friday,isSuite = False,
                          na = 0,nc = 0):
   total_reservations = len(hotelData_requests) 
   matching_reservations = 0
   numberOfRooms = 0;
   
   for index, col in hotelData_cancellations.iterrows():       
       if(col[2] == requested_room_type):
           if(getMonth(col[4]) == reservation_date_month):
              if(get_season(str(col[5])) == arrival_date_season):
                  if (col[8] == 'CC'):
                      if(col[7] == length_of_stay):
                          if(col[13] == weekend_friday):
                              if(isSuite == False):
                                  numberOfRooms+=1
                                  matching_reservations +=1
                              else:
                                  if((na==1 and nc ==0) or (na==1 and nc ==1) or (na==2 and nc ==0)):
                                      if((col[9]==1 and col[10] ==0) or (col[9]==1 and col[10] ==1) or (col[9]==2 and col[10] ==0)):
                                       numberOfRooms+=1
                                       matching_reservations +=1   
                                  elif((na==1 and nc ==2)or (na==2 and nc ==1)):
                                      if((col[9]==1 and col[10] ==2)or (col[9]==2 and col[10] ==1)):
                                        numberOfRooms+=1
                                        matching_reservations +=1  
                                  elif((na==1 and nc ==3) or (na==2 and nc ==2)):
                                      if((col[9]==1 and col[10] ==3) or (col[9]==2 and col[10] ==2)):
                                          numberOfRooms+=1
                                          matching_reservations +=1  
                                  elif((na==3 and nc ==0)):
                                      if((col[9]==3 and col[10] ==0)):
                                          numberOfRooms+=1
                                          matching_reservations +=1
                            
   
   # print(f"Room:{requested_room_type} reserved in: {reservation_date_month} for season {arrival_date_season} with length of stay of {length_of_stay} nights, On a weekend? {weekend_friday}, IsSuite? {isSuite}, Adults = {na} Children = {nc}")
   probability = matching_reservations / total_reservations
   # print(probability)
   return probability




# calculate_probability(hotelData,"Comfort Room","January",True,"LS",2)
# probability = calculate_probability(hotelData,"Comfort Room","June","VHS","CO",2,True)
# print("Probability:", probability)
# probability = calculate_prob_cancellations(hotelData_cancellations,"Comfort Room","January","LS",2,True,False,2,0)
# print("Probability:", probability)
probability = calculate_prob_cancellations(hotelData_cancellations,"Comfort Room","November","SS",1,True,True,2,0)
print("Probability:", probability)




# def getCancelledRooms(hotelData):
#     count = 0
#     for index, col in hotelData.iterrows():
#         if(col[8] == "CC" or col[8] == "CI" or col[8]== "EX" or col[8]== "NS"):
#             count+=1
#         if(col[8] == "CO" and col[7] >=4):
#             count+=1
#     return count