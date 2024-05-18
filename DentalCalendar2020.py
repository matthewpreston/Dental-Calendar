#!/usr/bin/env python3
# DentalCalendar.py - Designed to take events read from the official UofT
# Dentistry Excel file for clinical rotations and insert them into the
# official UofT Dentistry Microsoft Calendar file.
#
# Currently, when viewing the calendar, clinic times are written as just
# "Clinical Practice" as each person has an individualized clinic
# schedule found in the Excel file. They are to flip back and forth
# between these two to figure out where to go. This program is to merge
# them into one calendar format.
#
# Figuring out how to print the calendar into a pdf will be a future issue.
# Looks like one has to manually use the Outlook program (not the online one)
# in order to create a pdf.
#
### EDIT ###
# It appears that 2020 is a great year for us all and this program will not
# stand the test of time. With 3 clincs instead of 2 on Tues/Thurs, this year's
# program will have to written by hand and will truly be a "script"
#
# Written By: Matt Preston
# Written On: Aug 29, 2020
# Revised On: Nov  1, 2020 V2 - New calendar, new problems
#             Dec 20, 2020 V3 - Path of Pain

import argparse
from datetime import datetime, time, timedelta
import os
from pytz import timezone
import re

from icalendar import Calendar, Event   # For .ics files
from icalendar.prop import vText
from pandas import read_excel           # For Excel files

# Timezone data
EASTERN = timezone("Canada/Eastern")

# Days of week for datetime
MONDAY = 0
TUESDAY = 1
WEDNESDAY = 2
THURSDAY = 3
FRIDAY = 4
SATURDAY = 5
SUNDAY = 6

WEEKDAYS = {
    "Mon": 0,
    "Tue": 1,
    "Wed": 2,
    "Thu": 3,
    "Fri": 4,
    "Sat": 5,
    "Sun": 6
}

# Months to number
MONTHS = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12
}

# Number to months
JANUARY = 1
FEBRUARY = 2
MARCH = 3
APRIL = 4
MAY = 5
JUNE = 6
JULY = 7
AUGUST = 8
SEPTEMBER = 9
OCTOBER = 10
NOVEMBER = 11
DECEMBER = 12

# Unique ID generation for calendar events
UID = 0x040000008200E00074C5B7101A82E00800000000B018367A1691D3010000000000000000100000006AC9A0BE63E24944931F7635DF2D1C2E
UID_COUNTER = 0

# Custom colours for non-clinic events
NON_CLINIC_COLOUR_KEY = {
    "Lunch": vText("Yellow Category"),
    "Lunch ": vText("Yellow Category"),
    "Student Vendor Fair": vText("Green Category"),
    "Graduation Day": vText("Green Category"),
    "Winter Holidays ": vText("Green Category"),
    "Civic Holiday": vText("Green Category"),
    "CDSC Conference - Classes and Clinics Cancelled ": vText("Green Category"),
    "ODA Annual Spring Meeting ": vText("Green Category"),
    "Fall Study Day - Classes and Clinics Cancelled ": vText("Green Category"),
    "Canada Day - University Closed": vText("Green Category"),
    "Classes begin": vText("Green Category"),
    "Reading Week": vText("Green Category"),
    "Clinics close": vText("Green Category"),
    "Summer session begins": vText("Green Category"),
    "Classes end": vText("Green Category"),
    "Classes end and clinics close": vText("Green Category"),
    "Orientation Week": vText("Green Category"),
    "Thanksgiving Day": vText("Green Category"),
    "Labour day": vText("Green Category"),
    "Family Day": vText("Green Category"),
    "Victoria Day": vText("Green Category"),
    "Good Friday": vText("Green Category"),
    "Oral Medicine and Pathology - Seminars (14) - DEN315Y1": vText("Purple Category"),
    "IPC Distribution and Information ": vText("Red Category"),
    "Dental Public Health (DEN308Y) - Term Test": vText("Blue Category"),
    "Psychiatry and Dentistry - Test": vText("Blue Category"),
    "InterProfessional Pain Curriculum ": vText("Red Category"),
    "Oral Diagnosis & Medicine (DEN356Y) - Term Test": vText("Blue Category"),
    "Practice Administration (DEN409Y) - Term Test": vText("Blue Category"),
    "Oral Surgery (DEN318Y) - Term Test": vText("Blue Category"),
    "Pediatric Dentistry (DEN323Y) - Term Test": vText("Blue Category"),
    "Oral Med & Pathology (DEN315Y) ": vText("Blue Category"),
    "Prosthodontics (DEN333Y) - Term Test": vText("Blue Category"),
    "Oral Radiology (DEN317Y) - Term Test": vText("Blue Category"),
    "Anesthesia (DEN301Y) - Term Test": vText("Blue Category"),
    "Orthodontics (DEN322Y) - Term Test": vText("Blue Category"),
    "Endodontics (DEN303H) - Term Test": vText("Blue Category"),
    "Pharmacology (DEN327H) - Term Test": vText("Blue Category"),
    "DEN322Y1 Orthodontic - Final Exam": vText("Blue Category"),
    "DEN333Y1 Prosthodontics - Final Exam": vText("Blue Category"),
    "DEN315Y1 Oral Medicine & Pathology - Final Exam": vText("Blue Category"),
    "DEN323Y1 Pediatric Dentistry - Final Exam": vText("Blue Category"),
    "DEN336Y1 Restorative Dentistry - Final Exam": vText("Blue Category"),
    "Oral Radiology - part 1 - Final Exam": vText("Blue Category"),
    "Oral Radiology - part 2 - Final Exam": vText("Blue Category"),
    "Oral Radiology - Part 1 & 2 - Final Exam": vText("Blue Category"),
    "Anesthesia (DEN301H1) - Final Exam": vText("Blue Category"),
    "Pharmacology (DEN327H1) - Final Exam": vText("Blue Category"),
    "Endodontics (DEN303H1) - Final Exam": vText("Blue Category"),
    "Final Exam Period": vText("Blue Category"),
    "Periodontics (DEN324H) - Term Test": vText("Blue Category"),
    "Restorative (DEN336Y) - Term Test": vText("Blue Category"),
    "InterProfessional Pain Curriculum (1)": vText("Red Category"),
    "Research Day": vText("Green Category"),
    "DEN318Y1 Oral and Maxillofacial Surgery - Final Exam ": vText("Blue Category"),
    "Comprehensive Care - DEN451Y - Term Test": vText("Blue Category"),
    "Orthodontics - DEN465Y1 - Term Test": vText("Blue Category"),
    "Oral Radiology - DEN459Y1 - Term Test": vText("Blue Category"),
    "Oral Surgery - DEN462Y1 - Final Exam": vText("Blue Category"),
    "Pediatric Dentistry - DEN468Y - Term Test": vText("Blue Category"),
    "Comprehensive Care - DEN451Y - Test": vText("Blue Category"),
    "Practice Administration - Test": vText("Blue Category"),
    "Endodontics - DEN453Y1 - Term Test": vText("Blue Category"),
    "Anesthesia - DEN400H - Final Exam": vText("Blue Category"),
    "Dean's Welcome Back!": vText("Purple Category"),
    "IPE Orientation": vText("Purple Category"),
    "Orthodontics Screening": vText("Orange Category"),
    "Labour Day (University Closed)": vText("Green Category"),
    "Thanksgiving Day (University Closed)": vText("Green Category"),
    "Fall Study Day": vText("Green Category"),
    "Winter Holidays (University Closed)": vText("Green Category"),
    "Family Day (University Closed)": vText("Green Category"),
    "Study Day for NDEB Examinations": vText("Green Category"),
    "Good Friday (University Closed)": vText("Green Category"),
    "Oral Examination Period": vText("Blue Category")
}

class CheckFileAction(argparse.Action):
    """Ensures that file exists and is OK to read/write"""
    
    def __call__(self, parser, namespace, values, option_string=None):
        # It is actually better to open a file than it is to check existance,
        # else it could lead to annoying bugs (Why can't it find my file?!?!?
        # It's clearly right there you stupid program!)
        try:
            handle = open(values)
        except IOError as e:
            parser.error(e)
        else:
            handle.close()
        setattr(namespace, self.dest, values)

class CheckModeAction(argparse.Action):
    """Ensures that mode is in proper format"""
    
    def __call__(self, parser, namespace, values, option_string=None):
        if values not in ["All", "Clinics"]:
            parser.error("Invalid mode: {}; expected either 'All' or 'Clinics'"
                            "".format(values))
        setattr(namespace, self.dest, values)

def getStartTime(month, day, time, weekday, studentClinicID, clinicKey):
    """Returns start time for a given clinical date
    
    month(int)
        Month of session
    day(int)
        Day of session
    time("AM","PM","PM1","PM2")
        Time of session
    weekday(MONDAY,TUESDAY,WEDNESDAY,THURSDAY,FRIDAY)
        Which weekday
    studentClinicID(int)
        Which student
    clinicKey(str)
        Which clinic
        
    Returns (int, int)
        (Hour, minute) of session
    """
    
    # See if it's an ortho screening day and handle it
    if month == FEBRUARY and 15 <= day and day <= 26:
        if studentClinicID <= 30:
            if weekday == TUESDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
            if weekday == THURSDAY:
                if time == "AM":
                    return (8, 30)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
        if 30 < studentClinicID and studentClinicID <= 60:
            if weekday == TUESDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
            if weekday == THURSDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (12, 30)
                else:
                    return (16, 30)
        if 60 < studentClinicID and studentClinicID <= 90:
            if weekday == TUESDAY:
                if time == "AM":
                    return (8, 30)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
            if weekday == THURSDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
        if 90 < studentClinicID and studentClinicID <= 120:
            if weekday == TUESDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (12, 30)
                else:
                    return (16, 30)
            if weekday == THURSDAY:
                if time == "AM":
                    return (9, 00)
                elif time == "PM1":
                    return (13, 00)
                else:
                    return (16, 30)
    
    # See if it's end of the year and handle it
    if month == MAY:
        if 3 <= day and day <= 4:
            if studentClinicID <= 60:
                if weekday == MONDAY:
                    if time == "AM":
                        return (9, 00)
                    else:
                        return (13, 00)
                else: # Tue
                    if time == "AM":
                        return (8, 30)
                    else:
                        return (12, 30)
            else:
                if weekday == MONDAY:
                    if time == "AM":
                        return (8, 30)
                    elif time == "PM1":
                        return (12, 30)
                    else:
                        return (16, 30)
                else: # Tue
                    if time == "AM":
                        return (9, 00)
                    else:
                        return (13, 00)
        if day == 5:
            if time == "AM":
                return (9, 00)
            elif time == "PM1":
                return (13, 00)
            else:
                return (16, 30)
        if 6 <= day and day <= 14:
            return (9, 00) if time == "AM" else (13, 00)
    
    # Else handle normally
    if studentClinicID <= 60:
        if weekday == MONDAY:
            if time == "AM":
                return (9, 00)
            elif time == "PM" or time == "PM1":
                return (13, 00)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (16, 30)
                else:
                    return (16, 30)
        elif weekday == TUESDAY:
            if time == "AM":
                return (8, 00)
            elif time == "PM" or time == "PM1":
                return (11, 30)
            elif time == "PM2":       # 3rd clinic of day
                return (15, 00)
        elif weekday == WEDNESDAY:
            if time == "AM":
                return (9, 00)
            elif time == "PM" or time == "PM1":
                return (12, 30)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (16, 30)
                else:
                    return (16, 30)
        elif weekday == THURSDAY:
            if time == "AM":
                return (9, 00)
            elif time == "PM" or time == "PM1":
                return (13, 00)
            elif time == "PM2":       # Shouldn't have a hospital rotation, but in case I misread
                return (16, 30)
        elif weekday == FRIDAY:
            if time == "AM":
                return (8, 30)
            elif time == "PM" or time == "PM1":
                return (12, 30)
            elif time == "PM2":       # AGP session
                return (16, 30)
    else:
        if weekday == MONDAY:
            if time == "AM":
                return (8, 30)
            elif time == "PM" or time == "PM1":
                return (12, 30)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (16, 30)
                else:
                    return (16, 30)
        elif weekday == TUESDAY:
            if time == "AM":
                return (9, 00)
            elif time == "PM" or time == "PM1":
                return (13, 00)
            elif time == "PM2":       # Shouldn't have a hospital rotation, but in case I misread
                return (16, 30)
        elif weekday == WEDNESDAY:
            if time == "AM":
                return (8, 30)
            elif time == "PM" or time == "PM1":
                return (13, 00)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (16, 30)
                else:
                    return (16, 30)
        elif weekday == THURSDAY:
            if time == "AM":
                return (8, 00)
            elif time == "PM" or time == "PM1":
                return (11, 30)
            elif time == "PM2":       # 3rd clinic of day
                return (15, 00)
        elif weekday == FRIDAY:
            if time == "AM":
                return (9, 00)
            elif time == "PM" or time == "PM1":
                return (13, 00)
            elif time == "PM2":       # AGP session
                return (16, 30)

def getEndTime(month, day, time, weekday, studentClinicID, clinicKey):
    """Returns end time for a given clinical date
    
    month(int)
        Month of session
    day(int)
        Day of session
    time("AM","PM","PM1","PM2")
        Time of session
    weekday(MONDAY,TUESDAY,WEDNESDAY,THURSDAY,FRIDAY)
        Which weekday
    studentClinicID(int)
        Which student
    clinicKey(str)
        Which clinic
        
    Returns (int, int)
        (Hour, minute) of session
    """
    
    # See if it's an ortho screening day and handle it
    if month == FEBRUARY and 15 <= day and day <= 26:
        if studentClinicID <= 30:
            if weekday == TUESDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
            if weekday == THURSDAY:
                if time == "AM":
                    return (11, 30)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
        if 30 < studentClinicID and studentClinicID <= 60:
            if weekday == TUESDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
            if weekday == THURSDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (15, 30)
                else:
                    return (19, 00)
        if 60 < studentClinicID and studentClinicID <= 90:
            if weekday == TUESDAY:
                if time == "AM":
                    return (11, 30)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
            if weekday == THURSDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
        if 90 < studentClinicID and studentClinicID <= 120:
            if weekday == TUESDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (15, 30)
                else:
                    return (19, 00)
            if weekday == THURSDAY:
                if time == "AM":
                    return (12, 00)
                elif time == "PM1":
                    return (16, 00)
                else:
                    return (19, 00)
    
    # See if it's end of the year and handle it
    if month == MAY:
        if 3 <= day and day <= 4:
            if studentClinicID <= 60:
                if weekday == MONDAY:
                    if time == "AM":
                        return (12, 00)
                    else:
                        return (16, 00)
                else: # Tue
                    if time == "AM":
                        return (11, 30)
                    else:
                        return (15, 30)
            else:
                if weekday == MONDAY:
                    if time == "AM":
                        return (11, 30)
                    elif time == "PM1":
                        return (15, 30)
                    else:
                        return (19, 00)
                else: # Tue
                    if time == "AM":
                        return (12, 00)
                    else:
                        return (16, 00)
        if day == 5:
            if time == "AM":
                return (12, 00)
            elif time == "PM1":
                return (16, 00)
            else:
                return (19, 00)
        if 6 <= day and day <= 14:
            return (12, 00) if time == "AM" else (16, 00)
    
    # Else handle normally
    if studentClinicID <= 60:
        if weekday == MONDAY:
            if time == "AM":
                return (12, 00)
            elif time == "PM" or time == "PM1":
                return (16, 00)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (19, 30)
                else:
                    return (19, 00)
        elif weekday == TUESDAY:
            if time == "AM":
                return (10, 30)
            elif time == "PM" or time == "PM1":
                return (14, 00)
            elif time == "PM2":       # 3rd clinic of day
                return (17, 30)
        elif weekday == WEDNESDAY:
            if time == "AM":
                return (12, 00)
            elif time == "PM" or time == "PM1":
                return (15, 30)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30 or AGP
                if clinicKey == "PMH":
                    return (19, 30)
                else:
                    return (19, 00)
        elif weekday == THURSDAY:
            if time == "AM":
                return (12, 00)
            elif time == "PM" or time == "PM1":
                return (16, 00)
            elif time == "PM2":       # Shouldn't have a hospital rotation, but in case I misread
                return (19, 30)
        elif weekday == FRIDAY:
            if time == "AM":
                return (11, 30)
            elif time == "PM" or time == "PM1":
                return (15, 30)
            elif time == "PM2":       # AGP session
                return (19, 00)
    else:
        if weekday == MONDAY:
            if time == "AM":
                return (11, 30)
            elif time == "PM" or time == "PM1":
                return (15, 30)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30
                if clinicKey == "PMH":
                    return (19, 30)
                else:
                    return (19, 00)
        elif weekday == TUESDAY:
            if time == "AM":
                return (12, 00)
            elif time == "PM" or time == "PM1":
                return (16, 00)
            elif time == "PM2":       # Shouldn't have a hospital rotation, but in case I misread
                return (19, 30)
        elif weekday == WEDNESDAY:
            if time == "AM":
                return (11, 30)
            elif time == "PM" or time == "PM1":
                return (16, 00)
            elif time == "PM2":       # Possible hospital rotation from 4:30 - 7:30
                if clinicKey == "PMH":
                    return (19, 30)
                else:
                    return (19, 00)
        elif weekday == THURSDAY:
            if time == "AM":
                return (10, 30)
            elif time == "PM" or time == "PM1":
                return (14, 00)
            elif time == "PM2":       # 3rd clinic of day
                return (17, 30)
        elif weekday == FRIDAY:
            if time == "AM":
                return (12, 00)
            elif time == "PM" or time == "PM1":
                return (16, 00)
            elif time == "PM2":       # AGP session
                return (19, 00)

class Session:
    """Used to hold clinical session data.
    
    Members:
        clinic (str)
            What clinic you're supposed to be in
        room (str)
            What room the clinic is in
        description (str)
            Any details about the session
        start (datetime)
            Time clinic starts
        end (datetime)
            Time clinic ends
            
    Static members:
        CLINIC_KEY (dict(str: tuple(str, (lambda int, str), (lambda int, str))))
            Provides information about a particular clinic key string
            
        CLINIC_COLOUR_KEY(dict(str: icalendar.prop.vText))
            Creates an '''unique''' colour for each clinic key string
    
    Static methods:
        createSession(clinic, ID, start, end)
            Helper method to create a Session class
    """
    
    # Given a clinic symbol, provide what it means, room, and description
    CLINIC_KEY = {
        "FT": ("Faculty Timetable",          # Summary
                (lambda ID, start: "NA"),    # Room
                (lambda ID, start: "")),     # Description
        "C1": ("CCP - Clinic 1",
                (lambda ID, start: "Clinic 1"), 
                (lambda ID, start: "")),
        "C2": ("CCP - Clinic 2", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "CH": ("Pediatric Clinic", 
                (lambda ID, start: "Pediatric Clinic - 1st floor"), 
                (lambda ID, start: "")),
        "EM": ("Emergency Clinic", 
                (lambda ID, start: "Emergency Clinic - 2nd floor"), 
                (lambda ID, start: "")),
        "OD": ("Oral Diagnosis Clinic", 
                (lambda ID, start: "Oral Diagnosis Clinic - 2nd floor"), 
                (lambda ID, start: "")),
        "OR": ("Orthodontics Clinic", 
                (lambda ID, start: "481 University Avenue - 4th floor"), 
                (lambda ID, start: "")),
        "RA": ("Radiology Clinic", 
                (lambda ID, start: "Radiology Clinic - 2nd floor"), 
                (lambda ID, start: "")),
        "SC": ("Oral Surgery Clinic", 
                (lambda ID, start: "Oral Surgery Clinic - 1st floor"), 
                (lambda ID, start: "")),
        "HR": ("Hospital Rotations", 
                (lambda ID, start: "See Clinic Office Schedule"), # Done via hospitalFile parameter
                (lambda ID, start: "")),
        "GB": ("George Brown Health Sciences", 
                (lambda ID, start: "51 Dockside Drive (Take #6 bus on Bay St.)"),
                # Silly lambda's, have to do reverse if statemeent structure
                (lambda ID, start: \
                    "Arrive by 11:30 AM" \
                    if start.weekday() == MONDAY \
                    else "Arrive by 7:30 AM" \
                    if start.weekday() == TUESDAY \
                    else "")),
        "CA": ("CAMH Rotation", 
                (lambda ID, start: "100 Stokes Street - via Queen Street West"), 
                (lambda ID, start: \
                    "Placement begins at 9:15 AM" \
                    if start.time() < time(12) \
                    else "Placement begins at 1:15 AM")),
        "Ge": ("Assist in Grad Endo", 
                (lambda ID, start: "Grad Endo Clinics - 2nd floor"), 
                (lambda ID, start: "")),
        "Go": ("Assist in Oral Reconstruction", 
                (lambda ID, start: "Grad Perio/ORC Clinics - 3rd floor"), 
                (lambda ID, start: "")),
        "Gp": ("Assist in Grad Perio", 
                (lambda ID, start: "Grad Perio/ORC Clinics - 3rd floor"), 
                (lambda ID, start: \
                    # If PM Session:
                    "Arrive by 2:00 PM" \
                    if start.time() >= time(12) \
                    # If AM Session:
                    else \
                        # Then determine which weekday
                        "Arrive by 9:30 AM" \
                        if start.weekday() == MONDAY \
                        else "Arrive by 10:00 AM" \
                        if start.weekday() == TUESDAY \
                        else "Arrive by 8:30 AM" \
                        if start.weekday() == WEDNESDAY \
                        else "Arrive by 8:30 AM" \
                        if start.weekday() == THURSDAY \
                        else "Arrive by 10:00 AM")),
        "Gpr": ("Assist in Grad Prostho", 
                (lambda ID, start: "Grad Prostho Clinics - 3rd floor"), 
                (lambda ID, start: "")),
        "SM": ("St. Michael's Hospital Rotation", 
                (lambda ID, start: "80 Bond Street"), 
                (lambda ID, start: "Must register between 8:00-8:30 AM. Placement begins at 8:30 AM.")),
        "ST": ("Study Time", 
                (lambda ID, start: ""), 
                (lambda ID, start: "")),
        "ORL": ("Orthodontics Lab", 
                (lambda ID, start: "Senior Lab"), 
                (lambda ID, start: "")),
        "PS": ("Periodontic Suturing", 
                (lambda ID, start: "Lab 4"), 
                (lambda ID, start: "Reminder - Bring a banana to suture")),
        "PSC": ("Pediatric Surgicentre", 
                (lambda ID, start: "Adult Anaesthesia Clinic - Room 256"), 
                (lambda ID, start: "")),
        "CC": ("Restorative CAD/CAM Crowns", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "ENM": ("Endodontics of Molar", 
                (lambda ID, start: "Clinic 1"), 
                (lambda ID, start: "")),
        "R/1": ("Restorative Test 1 - Practice", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "R/2": ("Restorative Test 2 - Intracoronal", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "R/3": ("Restorative Test 3 - Crown", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "R/4": ("Restorative Test 4 - Anterior Resin", 
                (lambda ID, start: "Clinic 2"), 
                (lambda ID, start: "")),
        "AN3": ("Anaesthesia Clinic (Nitrous Oxide)", 
                (lambda ID, start: "Clinic 1"), 
                (lambda ID, start: "")),
        "IP": ("IPE - Pain", 
                (lambda ID, start: "External Placements - See Schedule from IPE Co-ordinator"), 
                (lambda ID, start: "")),
        "LT": ("IPE - Lindberg Homburger Modent Dental Studies Ltd.", 
                (lambda ID, start: "1407 Dufferin Street, Toronto, ON  M6H 4C7"), 
                (lambda ID, start: "")),
        "TPS": ("IPE - Toronto Paramedic Services", 
                (lambda ID, start: "4330 Dufferin St, North York, ON M3H 5R9"), 
                (lambda ID, start: "")),
        "PB": ("Patient Based Learning Seminar", 
                (lambda ID, start: "See Course Syllabus"), # TODO: Generate per student
                (lambda ID, start: "")),
        "P/F": ("Prosthodontics - Fixed Prostho Seminar", 
                (lambda ID, start: "See Course Syllabus"), # TODO
                (lambda ID, start: "")),
        "P1": ("Preventative Seminar #1", 
                (lambda ID, start: "See Course Syllabus"), # TODO
                (lambda ID, start: "")),
        "P2": ("Preventative Seminar #2", 
                (lambda ID, start: "See Course Syllabus"), # TODO
                (lambda ID, start: "")),
        "ORC": ("Oral Reconstruction Clinic",
                (lambda ID, start: "See Course Syllabus"), # TODO
                (lambda ID, start: "")),
        "GB": ("George Brown Health Sciences",
                (lambda ID, start: "51 Dockside Dr. - (#6 Bus South on Bay St.)"), # TODO
                (lambda ID, start: "")),
        "CA": ("CAMH Rotation",
                (lambda ID, start: "100 Stokes St."),
                (lambda ID, start: "Begins at 9:15 AM and 1:15 PM")),
        "SM": ("St. Michael's Hospital Rotation",
                (lambda ID, start: "80 Bond St."),
                (lambda ID, start: "Register at 8 AM. Placement begins at 8:30 AM")),
        "PMH": ("Princess Margaret Hospital Rotation",
                (lambda ID, start: "610 University Ave"),
                (lambda ID, start: "Placement from 4:30 PM - 7:30 PM")),
        "PB": ("Patient Based Learning Seminar", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "ET": ("Ethics Seminar", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "RS": ("Radiology Seminar", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "AN4": ("Anaesthesia Seminar", 
                (lambda ID, start: "Online - Synchronous, Room 360"), 
                (lambda ID, start: "")),
        "ORS": ("Orthodontics Seminar", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "OX": ("Orthodontics Oral Exam", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "CHS": ("Pediatric Clinic Seminar", 
                (lambda ID, start: "Online - Synchronous"), 
                (lambda ID, start: "")),
        "MS": ("Mount Sinai Rotation", 
                (lambda ID, start: "600 University Ave - 4th floor"), 
                (lambda ID, start: "")),
        "AG": ("AGP", 
                (lambda ID, start: "Closed Operatory"), 
                (lambda ID, start: "")),
        "AG-A": ("AGP Assisting", 
                (lambda ID, start: "Closed Operatory"), 
                (lambda ID, start: "")),
        "IPE": ("Grad Oral Surgery IPE", 
                (lambda ID, start: "Oral Surgery Clinic - 1st floor"), 
                (lambda ID, start: "Arrive by 8:45 AM" if start.time() >= time(12) else "Arrive by 12:45 PM"))
    }
    
    CLINIC_COLOUR_KEY = {
        "FT": vText("Lectures"),
        "C1": vText("Orange Category"),
        "C2": vText("Orange Category"),
        "CH": vText("Orange Category"),
        "EM": vText("Orange Category"),
        "OD": vText("Orange Category"),
        "OR": vText("Orange Category"),
        "RA": vText("Orange Category"),
        "SC": vText("Orange Category"),
        "HR": vText("Red Category"),
        "GB": vText("Red Category"),
        "CA": vText("Red Category"),
        "Ge": vText("Red Category"),
        "Go": vText("Red Category"),
        "Gp": vText("Red Category"),
        "Gpr": vText("Red Category"),
        "SM": vText("Red Category"),
        "ST": vText("Green Category"),
        "ORL": vText("Purple Category"),
        "PS": vText("Purple Category"),
        "PSC": vText("Red Category"),
        "CC": vText("Red Category"),
        "ENM": vText("Purple Category"),
        "R/1": vText("Blue Category"),
        "R/2": vText("Blue Category"),
        "R/3": vText("Blue Category"),
        "R/4": vText("Blue Category"),
        "AN3": vText("Orange Category"),
        "IP": vText("Red Category"),
        "LT": vText("Red Category"),
        "TPS": vText("Red Category"),
        "PB": vText("Purple Category"),
        "P/F": vText("Purple Category"),
        "P1": vText("Purple Category"),
        "P2": vText("Purple Category"),
        "ORC": vText("Red Category"),
        "GB": vText("Red Category"),
        "CA": vText("Red Category"),
        "SM": vText("Red Category"),
        "PMH": vText("Red Category"),
        "PB": vText("Purple Category"),
        "ET": vText("Purple Category"),
        "RS": vText("Purple Category"),
        "AN4": vText("Purple Category"),
        "ORS": vText("Purple Category"),
        "OX": vText("Blue Category"),
        "CHS": vText("Purple Category"),
        "MS": vText("Red Category"),
        "AG": vText("Red Category"),
        "AG-A": vText("Red Category"),
        "IPE": vText("Red Category")
    }
    
    def __init__(self, clinic, room, description, start, end, colour):
        self.clinic = clinic
        self.room = room
        self.description = description
        self.start = start
        self.end = end
        self.colour = colour

    @staticmethod
    def createSession(clinicKey, studentClinicID, start, end):
        """Helper method to create a Session class
        
        clinicKey (str)
            A string key used to determine which clinic the key refers to
        studentClinicID (int)
            Student clinical number
        start (datetime)
            When the session starts
        end (datetime)
            When the session ends
        returns Session
            Contains everything you need to know about a clinic timeslot
        """
        
        assert start.weekday() >= MONDAY and start.weekday() <= FRIDAY
        
        if clinicKey in Session.CLINIC_KEY:
            summary, roomFunc, descFunc = Session.CLINIC_KEY[clinicKey]
            colour = Session.CLINIC_COLOUR_KEY[clinicKey]
        else:
            summary = clinicKey
            roomFunc = lambda ID, start: ""
            descFunc = lambda ID, start: ""
            colour = vText("")
        room = roomFunc(studentClinicID, start)
        desc = descFunc(studentClinicID, start)
        return Session(summary, room, desc, start, end, colour)

def createDatetime(excelDataframe, excelRow, studentClinicID, clinicKey, timezone=EASTERN):
    """Given a row from the clinical Excel file, provide the start and end of
    that session.
    
    excelDataframe (pandas.dataframe)
        The Excel sheet in dataframe format
    excelRow (int)
        The row of the Excel file
    studentClinicID (int)
        ID of the student (now that since Mon/Thurs students have wildly
        different times than Tues/Fri students)
    clinicKey (str)
        ID of clinical session (see Session.CLINIC_KEY)
    timezone (datetime.tzinfo) [EASTERN]
        The desired timezone
    return (datetime, datetime)
        The start and end times of the session
    """
    
    # Excel file is indexed as follows (assume that we start on a Monday):
    # Row, Day, Date, Time, ...
    # n+0,  Mon,  9-Dec-19, AM, ...
    # n+1,  Mon,  9-Dec-19, PM1, ...
    # n+2,  Mon,  9-Dec-19, PM2, ...
    # n+3,    ,         ,   , ...
    # n+4,  Tue, 10-Dec-19, AM, ...
    # n+5,  Tue, 10-Dec-19, PM1, ...
    # n+6,  Tue, 10-Dec-19, PM2, ...
    # n+7,    ,         ,   , ...
    # n+8,  Wed, 11-Dec-19, AM, ...
    # n+9,  Wed, 11-Dec-19, PM1, ...
    # n+10, Wed, 11-Dec-19, PM2, ...
    # n+11,    ,         ,   , ...
    # n+12, Thu, 12-Dec-19, AM, ...
    # n+13, Thu, 12-Dec-19, PM1, ...
    # n+14, Thu, 12-Dec-19, PM2, ...
    # n+15,    ,         ,   , ...
    # n+16, Fri, 13-Dec-19, AM, ...
    # n+17, Fri, 13-Dec-19, PM, ...
    # ...
    # We want to steal the columns 'Day', 'Date', and 'Time' which are 'Unnamed 0',
    # 'Unnamed 1' and 'Unnamed 2' in the dataframe object
    
    weekday = WEEKDAYS[excelDataframe.at[excelRow, "Unnamed: 0"]]
    date = excelDataframe.at[excelRow, "Unnamed: 1"]
    time = excelDataframe.at[excelRow, "Unnamed: 2"]
    
    # Extract date data
    year = date.year
    month = date.month
    day = date.day
    
    # Extract time data
    startHour, startMinute = getStartTime(month, day, time, weekday, studentClinicID, clinicKey)
    start = timezone.localize(datetime(year, month, day, startHour, startMinute, 0))
    endHour, endMinute = getEndTime(month, day, time, weekday, studentClinicID, clinicKey)
    end = timezone.localize(datetime(year, month, day, endHour, endMinute, 0))
    return (start, end)

def standardizeDatetime(dateTime, timezone=EASTERN):
    """To get the datetime objects from the calendar file to comply with my
    manually set up ones.
    
    dateTime (datetime)
        The calendar's datetime
    timezone (datetime.tzinfo) [EASTERN]
        The desired timezone
    returns (datetime)
        A proper datetime
    """
    year = dateTime.year
    month = dateTime.month
    day = dateTime.day
    hour = dateTime.hour
    minute = dateTime.minute
    second = dateTime.second
    return timezone.localize(datetime(year, month, day, hour, minute, second))
    
def fixDatetime(dateTime, numWeeks, timezone=EASTERN):
    """Advances the provided datetime 'numWeeks' ahead, accounting for daylight
    saving time.
    
    dateTime (datetime)
        Provided datetime to advance
    numWeeks (int)
        Number of weeks to advance
    timezone (datetime.tzinfo) [EASTERN]
        The desired timezone
    returns (datetime)
        The advanced datetime
    """
    
    # See if given datetime is in DST
    wasDST = dateTime.dst() != timedelta(0)
    
    # If after advancing by 'numWeeks', we've now changed timezones
    tempDT = dateTime + timedelta(weeks=numWeeks)
    if (timezone.normalize(tempDT).dst() != timedelta(0)) ^ wasDST:
        if wasDST:  # Was DST, now is Standard Time
            tempDT += dateTime.dst()
        else:       # Was STD, now DST
            tempDT -= timezone.normalize(tempDT).dst()
    
    return timezone.normalize(tempDT)

def main(args):
    clinicFile = args.clinicFile
    calendarFile = args.calendarFile
    outputDir = args.outputDir
    mode = args.mode
    startStudentID = int(args.start)
    endStudentID = int(args.end)
    
    # Go through calendar, gathering events
    with open(calendarFile, "rb") as calendarHandle:
        cal = Calendar.from_ical(calendarHandle.read())
        components = list(cal.walk())
    
    # Go through clinical schedule, gathering which clinics people are in at
    # whatever dates and times
    clinics = read_excel(clinicFile, sheet_name=0)
    
    # Magic sequence that indicates the start of each session
    startOfWeeks = []
    # For Sept to Dec
    startOfWeeks.append([59 - 2 + i * 29 + (i+1)//2 for i in range(15)])
    # Then for Jan - NDEB study break (all 5 days have AM, PM1, PM2)
    startOfWeeks.append([553 - 2 + i * 30 + i//2 for i in range(8)])
    # Then NDEB study break to May
    startOfWeeks.append([822 - 2 + i * 30 + (i+1)//2 for i in range(9)])
    # Finally last week of May (only AM and PM)
    startOfWeeks.append([1097 - 2])
    
    # Magic sequence for each AM, PM1, PM2, etc. for Mon-Fri based on the above splits
    individualSessions = []
    individualSessions.append([0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17])
    individualSessions.append([0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18])
    individualSessions.append([0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18])
    individualSessions.append([0, 1, 3, 4, 6, 7, 9, 10, 12, 13])
    
    # Final magic sequence for Excel rows containing clinical sessions
    sessions = []
    for (weekBlocks, iSessions) in zip(startOfWeeks, individualSessions):
        for i in weekBlocks:
            for j in iSessions:
                sessions.append(i+j)
    
    # Magic sequence for columns
    clinicNumberCols = ["Section"] \
                     + ["Unnamed: {}".format(i) for i in range(4,23)] \
                     + ["Section.1"] \
                     + ["Unnamed: {}".format(i) for i in range(24,43)] \
                     + ["Section.2"] \
                     + ["Unnamed: {}".format(i) for i in range(44,63)] \
                     + ["Section.3"] \
                     + ["Unnamed: {}".format(i) for i in range(67,86)] \
                     + ["Section.4"] \
                     + ["Unnamed: {}".format(i) for i in range(87,106)] \
                     + ["Section.5"] \
                     + ["Unnamed: {}".format(i) for i in range(107,126)]
    
    # Now to finally parse the Excel file and extract which clinic should
    # someone be at what time
    clinicData = []
    for (studentClinicID, col) in enumerate(clinicNumberCols):
        studentClinicID += 1 # Make it indexed starting at 1
        clinicData.append(dict())
        for sess in sessions:
            clinicKey = str(clinics.at[sess, col])
            start, end = createDatetime(clinics, sess, studentClinicID, clinicKey)
            newSession = Session.createSession(clinicKey, studentClinicID, start, end)
            clinicData[studentClinicID-1][start] = newSession

    # Create the output directory
    if not os.path.exists(outputDir):
        os.mkdir(outputDir)
        
    # Create a calendar for each student
    for studentClinicID in range(startStudentID, endStudentID+1): #TODO testing
        # Skip non-existing students
        if studentClinicID == 61 or studentClinicID == 120:
            continue
    
        UID_COUNTER = 0
        newCal = Calendar()
        newCal.add("prodid", cal.get("prodid"))
        newCal.add("version", cal.get("version"))
        for c in components:
            if c.name == "VEVENT":
                # Find events that are clinical sessions and update them
                if "Clinical Practice" in str(c.get("summary")) or "Ancillary Clinics" in str(c.get("summary")):
                    eStart = standardizeDatetime(c.get("dtstart").dt)
                    
                    # These events are programmed to occur every week, skipping
                    # some when noted. So, create a series of new events for
                    # each valid week.
                    try:
                        repeats = c.get("rrule")["COUNT"][0]
                    except:
                        repeats = 1
                    try:
                        skips = [dt.dt for dt in c.get("exdate").dts]
                    except:
                        skips = []
                    for r in range(repeats):
                        # Skip 'r' weeks from the starting point
                        tempdt = fixDatetime(eStart, r)
                        
                        # If this date is to be skipped (holiday, hospital etc.)
                        if tempdt in skips:
                            continue
                        session = clinicData[studentClinicID-1][tempdt]
                        
                        # If this is a PM2 session for AGP and the Excel file
                        # says that it's a study time or faculty time, skip it
                        if tempdt.hour == 16 and tempdt.minute == 30 \
                                and (session.clinic == Session.CLINIC_KEY["ST"][0] \
                                     or session.clinic == Session.CLINIC_KEY["FT"][0]):
                            continue
                            
                        # Create a new event based on this time and add a bunch
                        # of junk to make the calendar uptake it
                        event = Event()
                        event.add("categories", session.colour)
                        event.add("class", c.get("class"))
                        event.add("created", c.get("created"))
                        event.add("dtstart", session.start)
                        event.add("dtend", session.end)
                        event.add("dtstamp", c.get("dtstamp"))
                        event.add("description", session.description)
                        event.add("last-modified", c.get("last-modified"))
                        event.add("location", session.room)
                        event.add("priority", c.get("priority"))
                        event.add("sequence", c.get("sequence"))
                        event.add("summary", session.clinic)
                        event.add("transp", c.get("transp"))
                        event.add("UID", "{:X}".format(UID + UID_COUNTER))
                        UID_COUNTER += 1
                        newCal.add_component(event)
                else: # Intercept it and change its colour
                    if mode == "Clinics": # If only clinics are to be outputted
                        continue
                        
                    event = Event()
                    for k in c.keys():
                        k = k.lower()
                        if k == "categories":
                            summary = str(c.get("summary"))
                            if summary in NON_CLINIC_COLOUR_KEY.keys():
                                event.add("categories", NON_CLINIC_COLOUR_KEY[summary])
                            else:
                                event.add("categories", c.get("categories"))
                        elif k == "uid":
                            event.add("UID", "{:X}".format(UID + UID_COUNTER))
                            UID_COUNTER += 1
                        elif k == "x-alt-desc" \
                                or k == "x-microsoft-cdo-busystatus" \
                                or k == "x-microsoft-cdo-importance" \
                                or k == "x-microsoft-disallow-counter":
                            pass
                        else:
                            event.add(k, c.get(k))
                    newCal.add_component(event)
        
        # Write calendar to file
        outputFile = "{}/{} - {}.ics".format(outputDir,
            calendarFile.split(".",1)[0][:25], studentClinicID)
        with open(outputFile, "wb") as output:
            output.write(newCal.to_ical())
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    
    # Add positional arguments
    parser.add_argument("clinicFile",
                        metavar="<clinic.xlsx>",
                        action=CheckFileAction,
                        help="""Excel file containing clinic data""")
    parser.add_argument("calendarFile",
                        metavar="<calendar.ics>",
                        action=CheckFileAction,
                        help="""Microsoft Calendar .ics file containing class
                            schedule""")
                            
    # Add optional arguments
    parser.add_argument("-o", "--outputDir",
                        metavar="DIR",
                        default="Dental Calendars",
                        help="""Output directory to store new .ics files
                            [Dental Calendars/]""")
    parser.add_argument("-s", "--start",
                        metavar="int",
                        default=1,
                        help="""Starting student clinic ID number""")
    parser.add_argument("-e", "--end",
                        metavar="int",
                        default=120,
                        help="""Ending student clinic ID number""")
    parser.add_argument("-m", "--mode",
                        metavar="[All,Clinics]",
                        action=CheckModeAction,
                        default="All",
                        help="""All: Full calendar generated. Clinics: Just
                            clinics generated""")
    args = parser.parse_args()
    main(args)
