import openpyxl

ACADEMIC_PERIOD_COL = 'A'
SECTION_COL = 'B'
MEETING_TIME_COL = 'C'
LOCATION_COL = 'D'
CAPACITY_COL = 'E'



class Event:
    def __init__(self, row):
        row_to_event(self, row)
        
    def row_to_event(self, row):
