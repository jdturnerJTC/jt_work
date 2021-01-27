##################################################################
# Scraping Data on Franklin County court records
# Author: CLH
# Date: 5-22-2020
##################################################################

# Step 1: Imports

import sys

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import datetime
from datetime import date, timedelta
import time
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import re
from selenium.webdriver import Chrome
from selenium.webdriver import ChromeOptions
import openpyxl
import xlrd

# Set current date and file to save -- CHANGE THIS

today = "9.10.20"
filesave = '/Volumes/leo/CSH Ohio/PII/Eligibility/Eligible 15-18 not in study scraped {}.csv'.format(
    today)
print(filesave)

# Set up Chrome Driver

opts = ChromeOptions()

opts.add_argument('headless')

# CHANGE THIS - change file path

chrome_location = "chromedriver"
browser = webdriver.Chrome(
    options=opts, executable_path="/Users/charles/Desktop/chromedriver")

# Go to Franklin County Site
url = "https://fcsojmsweb.franklincountyohio.gov/Publicview/(S(14knmnmsuo4boqb3t2l4aw0k))/BookingFind.aspx"
browser.get(url)
browser.refresh()

# Download excel file for eligible 15-18 - CHANGE THIS

wb = xlrd.open_workbook(
    r'/Volumes/leo/CSH Ohio/PII/Eligibility/Eligible people 15-18 not in study_7.30.20.xlsx')
sh = wb.sheet_by_name('Sheet1')

my_df = []

# search names in the file, save all data as variables, and append my file
# CHECK THIS - may need to change this if columns are in different orders

for i in range(1, sh.nrows):
    fn = str(sh.cell_value(i, 0))
    ln = str(sh.cell_value(i, 1))
    fn = fn.split()[0]
    FullName = fn + ln

    print(fn)
    print(ln)

    search_form1 = browser.find_element_by_id('InmateLast')
    search_form1.send_keys(ln)

    search_form2 = browser.find_element_by_id('InmateFirst')
    search_form2.send_keys(fn)

    search_form3 = browser.find_element_by_id('btnSearch')
    search_form3.click()

    try:
        details = browser.find_element_by_link_text('Detail')
        details.click()

        InmateID = browser.find_elements_by_id('InmateId')
        InmateName = browser.find_elements_by_id('InmateName')
        BookingID = browser.find_elements_by_id('BookingId')
        InmateAlias = browser.find_elements_by_id('InmateAKAName')
        GenderDesc = browser.find_elements_by_id('GenderDesc')
        RaceDesc = browser.find_elements_by_id('RaceDesc')
        Height = browser.find_elements_by_id('Height')
        Weight = browser.find_elements_by_id('Weight')
        HairColor = browser.find_elements_by_id('HairColor')
        BirthDate = browser.find_elements_by_id('BirthDate')
        PhotoDate = browser.find_elements_by_id('PhotoDate')
        BookingLocation = browser.find_elements_by_id('BookingLocation')
        InmateStatus = browser.find_elements_by_id('InmateStatus')
        CustodyLevel = browser.find_elements_by_id('CustodyLevel')
        EstimatedReleaseDate = browser.find_elements_by_id(
            'EstimatedReleaseDate')
        SID = browser.find_elements_by_id('SID')

        for value in InmateID:
            InmateID = value.get_attribute('value')

        for value in InmateName:
            InmateName = value.get_attribute('value')

        for value in BookingID:
            BookingID = value.get_attribute('value')

        for value in InmateAlias:
            InmateAlias = value.get_attribute('value')

        for value in GenderDesc:
            GenderDesc = value.get_attribute('value')

        for value in RaceDesc:
            RaceDesc = value.get_attribute('value')

        for value in Height:
            Height = value.get_attribute('value')

        for value in Weight:
            Weight = value.get_attribute('value')

        for value in HairColor:
            HairColor = value.get_attribute('value')

        for value in BirthDate:
            BirthDate = value.get_attribute('value')

        for value in PhotoDate:
            PhotoDate = value.get_attribute('value')

        for value in BookingLocation:
            BookingLocation = value.get_attribute('value')

        for value in InmateStatus:
            InmateStatus = value.get_attribute('value')

        for value in CustodyLevel:
            CustodyLevel = value.get_attribute('value')

        for value in EstimatedReleaseDate:
            EstimatedReleaseDate = value.get_attribute('value')

        for value in SID:
            SID = value.get_attribute('value')

        rows_of_table = browser.find_elements_by_xpath(
            '//*[@id="dgCases"]/tbody//tr[td]')

        rows = int(len(rows_of_table))
        x = rows + 1

        for c in range(2, x):
            offense = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[1]')
            case_no = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[2]')
            sentence = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[3]')
            convict_type = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[4]')
            date_of_arrest = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[5]')
            date_of_sentence = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[6]')
            bond_amt = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[7]')
            bond_type = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[8]')
            next_court_date = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[9]')
            court = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[10]')

            for j in offense:
                offense = j.text.replace('\n', '')
            for j in case_no:
                case_no = j.text.replace('\n', '')
            for j in sentence:
                sentence = j.text.replace('\n', '')
            for j in convict_type:
                convict_type = j.text.replace('\n', '')
            for j in date_of_arrest:
                date_of_arrest = j.text.replace('\n', '')
            for j in date_of_sentence:
                date_of_sentence = j.text.replace('\n', '')
            for j in bond_amt:
                bond_amt = j.text.replace('\n', '')
            for j in bond_type:
                bond_type = j.text.replace('\n', '')
            for j in next_court_date:
                next_court_date = j.text.replace('\n', '')
            for j in court:
                court = j.text.replace('\n', '')

            d = {'First Name': fn, 'Last Name': ln,
                 'Full Name': FullName,
                 'Inmate ID': InmateID, 'Inmate Name': InmateName, 'Booking ID': BookingID,
                 'Inmate Alias': InmateAlias, 'Gender': GenderDesc, 'Race': RaceDesc,
                 'Height': Height, 'Weight': Weight, 'Hair Color': HairColor,
                 'Birth Date': BirthDate, 'Photo Date': PhotoDate,
                 'Booking Location': BookingLocation, 'Inmate Status': InmateStatus,
                 'Custody Level': CustodyLevel, 'Estimated Release Date': EstimatedReleaseDate,
                 'SID': SID, 'Offense': offense, 'Case Number': case_no,
                 'Sentence': sentence, 'Conviction Type': convict_type,
                 'Date of Arrest': date_of_arrest, 'Date of Sentence': date_of_sentence,
                 'Bond Amount': bond_amt, 'Bond Type': bond_type,
                 'Next Court Date': next_court_date, 'Court': court}

            my_df.append(d)

        browser.back()
        browser.back()
        browser.refresh()

    except:

        browser.back()
        browser.refresh()


# save file as csv

my_df = pd.DataFrame(my_df)

print(my_df)

my_df.to_csv(filesave)

"""
# can repeat for almost eligible 16-18

url = "https://fcsojmsweb.franklincountyohio.gov/Publicview/(S(14knmnmsuo4boqb3t2l4aw0k))/BookingFind.aspx"
browser.get(url)
browser.refresh()


wb = xlrd.open_workbook(
    r'/Volumes/leo/CSH Ohio/PII/Eligibility/almosteligible_people_16-18.csv')
sh = wb.sheet_by_name('Sheet1')

my_df = []


for i in range(1, sh.nrows):
    ln = str(sh.cell_value(i, 0))
    fn = str(sh.cell_value(i, 1))
    fn = fn.split()[0]
    FullName = fn + ln

    print(fn)
    print(ln)

    search_form1 = browser.find_element_by_id('InmateLast')
    search_form1.send_keys(ln)

    search_form2 = browser.find_element_by_id('InmateFirst')
    search_form2.send_keys(fn)

    search_form3 = browser.find_element_by_id('btnSearch')
    search_form3.click()

    try:
        details = browser.find_element_by_link_text('Detail')
        details.click()

        InmateID = browser.find_elements_by_id('InmateId')
        InmateName = browser.find_elements_by_id('InmateName')
        BookingID = browser.find_elements_by_id('BookingId')
        InmateAlias = browser.find_elements_by_id('InmateAKAName')
        GenderDesc = browser.find_elements_by_id('GenderDesc')
        RaceDesc = browser.find_elements_by_id('RaceDesc')
        Height = browser.find_elements_by_id('Height')
        Weight = browser.find_elements_by_id('Weight')
        HairColor = browser.find_elements_by_id('HairColor')
        BirthDate = browser.find_elements_by_id('BirthDate')
        PhotoDate = browser.find_elements_by_id('PhotoDate')
        BookingLocation = browser.find_elements_by_id('BookingLocation')
        InmateStatus = browser.find_elements_by_id('InmateStatus')
        CustodyLevel = browser.find_elements_by_id('CustodyLevel')
        EstimatedReleaseDate = browser.find_elements_by_id(
            'EstimatedReleaseDate')
        SID = browser.find_elements_by_id('SID')

        for value in InmateID:
            InmateID = value.get_attribute('value')

        for value in InmateName:
            InmateName = value.get_attribute('value')

        for value in BookingID:
            BookingID = value.get_attribute('value')

        for value in InmateAlias:
            InmateAlias = value.get_attribute('value')

        for value in GenderDesc:
            GenderDesc = value.get_attribute('value')

        for value in RaceDesc:
            RaceDesc = value.get_attribute('value')

        for value in Height:
            Height = value.get_attribute('value')

        for value in Weight:
            Weight = value.get_attribute('value')

        for value in HairColor:
            HairColor = value.get_attribute('value')

        for value in BirthDate:
            BirthDate = value.get_attribute('value')

        for value in PhotoDate:
            PhotoDate = value.get_attribute('value')

        for value in BookingLocation:
            BookingLocation = value.get_attribute('value')

        for value in InmateStatus:
            InmateStatus = value.get_attribute('value')

        for value in CustodyLevel:
            CustodyLevel = value.get_attribute('value')

        for value in EstimatedReleaseDate:
            EstimatedReleaseDate = value.get_attribute('value')

        for value in SID:
            SID = value.get_attribute('value')

        rows_of_table = browser.find_elements_by_xpath(
            '//*[@id="dgCases"]/tbody//tr[td]')

        rows = int(len(rows_of_table))
        x = rows + 1

        for c in range(2, x):
            offense = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[1]')
            case_no = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[2]')
            sentence = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[3]')
            convict_type = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[4]')
            date_of_arrest = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[5]')
            date_of_sentence = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[6]')
            bond_amt = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[7]')
            bond_type = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[8]')
            next_court_date = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[9]')
            court = browser.find_elements_by_xpath(
                '//*[@id="dgCases"]/tbody/tr[' + str(c) + ']/td[10]')

            for j in offense:
                offense = j.text.replace('\n', '')
            for j in case_no:
                case_no = j.text.replace('\n', '')
            for j in sentence:
                sentence = j.text.replace('\n', '')
            for j in convict_type:
                convict_type = j.text.replace('\n', '')
            for j in date_of_arrest:
                date_of_arrest = j.text.replace('\n', '')
            for j in date_of_sentence:
                date_of_sentence = j.text.replace('\n', '')
            for j in bond_amt:
                bond_amt = j.text.replace('\n', '')
            for j in bond_type:
                bond_type = j.text.replace('\n', '')
            for j in next_court_date:
                next_court_date = j.text.replace('\n', '')
            for j in court:
                court = j.text.replace('\n', '')

            d = {'First Name': fn, 'Last Name': ln,
                 'Full Name': FullName,
                 'Inmate ID': InmateID, 'Inmate Name': InmateName, 'Booking ID': BookingID,
                 'Inmate Alias': InmateAlias, 'Gender': GenderDesc, 'Race': RaceDesc,
                 'Height': Height, 'Weight': Weight, 'Hair Color': HairColor,
                 'Birth Date': BirthDate, 'Photo Date': PhotoDate,
                 'Booking Location': BookingLocation, 'Inmate Status': InmateStatus,
                 'Custody Level': CustodyLevel, 'Estimated Release Date': EstimatedReleaseDate,
                 'SID': SID, 'Offense': offense, 'Case Number': case_no,
                 'Sentence': sentence, 'Conviction Type': convict_type,
                 'Date of Arrest': date_of_arrest, 'Date of Sentence': date_of_sentence,
                 'Bond Amount': bond_amt, 'Bond Type': bond_type,
                 'Next Court Date': next_court_date, 'Court': court}

            my_df.append(d)

        browser.back()
        browser.back()
        browser.refresh()

    except:

            browser.back()
            browser.refresh()


my_df = pd.DataFrame(my_df)

print(my_df)

my_df.to_csv(
    '/Volumes/leo/CSH Ohio/PII/Eligibility/Almost Eligible16-18 scraped.csv')"""
