from datetime import datetime



def GetLowRate(RoomType):
    switcher={
        1:90,
        2:105,
        3:125,
        4:140,
        5:150,
        6:165,
        7:200,
        8:230,
        }
    
    return switcher.get(RoomType,"nothing")

def GetShoulderRate(RoomType):
    switcher={
        1:105,
        2:120,
        3:140,
        4:155,
        5:165,
        6:180,
        7:215,
        8:240,
        }
    
    return switcher.get(RoomType,"nothing")

def GetHighRate(RoomType):
    switcher={
        1:120,
        2:135,
        3:155,
        4:170,
        5:180,
        6:195,
        7:225,
        8:250,
        }
    
    return switcher.get(RoomType,"nothing")

def GetVeryHighRate(RoomType):
    switcher={
        1:130,
        2:145,
        3:160,
        4:175,
        5:185,
        6:200,
        7:235,
        8:260,
        }
    
    return switcher.get(RoomType,"nothing")


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

def get_season(arrival_date):
    arrival_date = datetime.strptime(arrival_date, "%Y/%m/%d")
    
    for season, date_ranges in seasons.items():
        for start, end in date_ranges:
            start_date = datetime.strptime(start, "%Y-%m-%d")
            end_date = datetime.strptime(end, "%Y-%m-%d")
            if start_date <= arrival_date <= end_date:
                return season
    return "Invalid Date"
        
        
def GetRev(k):
    na= int(k.na);
    nc = int(k.nc);
    c=int(k.getRoomType());
    l = int(k.length);
    numberpersons = na+nc;
    
    season = get_season(k.ad)
    
    if season == "VHS":
        BaseRate = GetVeryHighRate(c)
    elif season == "HS":
        BaseRate = GetHighRate(c)
    elif season == "LS":
        BaseRate= GetLowRate(c)
    elif season == "SS":
        BaseRate = GetShoulderRate(c)
        
    if numberpersons <=2:
        return l*BaseRate;
    elif numberpersons > 2:
        if na == 3:
            return (l*BaseRate)+((0.75*(0.5*BaseRate))*l);
        elif na == 2:
            return (l*BaseRate)+(l*((0.5*(0.5*BaseRate))*nc));
        elif na == 1:
            return (l*BaseRate)+((0.5*(0.5*BaseRate))*(numberpersons-2)*l);