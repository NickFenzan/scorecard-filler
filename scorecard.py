#! /usr/bin/python3
import argparse
from scorecardfillerfunctions import *

parser = argparse.ArgumentParser(description='Fill numbers into scorecard')
parser.add_argument('scorecard', metavar='s', type=str, help='the existing scorecard')
parser.add_argument('appointments', metavar='a', type=str, help='appointments report')
parser.add_argument('billed', metavar='b', type=str, help='billed report')
parser.add_argument('collected', metavar='c', type=str, help='collected report')
parser.add_argument('staffing', metavar='staff', type=str, help='staffing report')
parser.add_argument('websiteform', metavar='w', type=str, help='website forms report')
parser.add_argument('calls', metavar='calls', type=int, help='calls')
parser.add_argument('websessions', metavar='sess', type=int, help='sessions')
args = parser.parse_args()

sites = ['Novi', 'Troy', 'Dearborn', 'Monroe', 'Macomb']
apptCategories = {
    'Free Consults': ['Free Evaluation'],
    'New Patients': ['New Patient'],
    'Procedures': ['EVCA', 'EVTA', 'EVTA MP', 'Microphlebectomy', 'MOCA', 'Varithena', 'VenaSeal', 'VVEVCA'],
    'Long Term Follow Ups': ['3 Month Follow Up', '6 Month Follow Up', 'Yearly Follow Up'],
    'VeinErase': ['VeinErase Legs', 'VeinErase MESSA', 'VeinErase - Facial']
}
staffRoles = ['Medtech', 'Ultrasound', 'Nurse', 'Physician']

writeScorecard(sites, apptCategories, staffRoles, args)
