# Mediation Review Assistant Deluxe, Python Version
# Written by Danny Brown -- Written January 29 through February 6, 2018
# Major revisions March 2 through 6, 2018 and early April 2018

import os
import string
import datetime
import win32com.client as win32
import modules.validate as validate
import modules.correspondence as cor
from nameparser import HumanName
from string import digits


#TODO collect reasons for mediation with option for none; save to data file
#TODO save Versa entry data and look up for subsequent letters


NAME = "Danny Brown"
TITLE = "Investigator Specialist II"
EMAIL = "daniel.brown@myfloridalicense.com"
PHONE = "850-487-9948"


def main():
    print "\nPlease select one of the following program functions:\n"
    print "1. Get assistance with reviewing a petition for mediation."
    print "2. Select the next mediator from rotation for an approved petition."
    print "3. Create letters from templates for mailing."
    print "4. Determine the judicial circuit for a city."
    userChoice = validate.menu("What would you like to do?", "1234")
    if userChoice is "1":
        petition_review()
    elif userChoice is "2":
        select_mediator()
    elif userChoice is "3":
        letter_center("", "", "", "")
    elif userChoice is "4":
        determine_judicial_circuit()
    exit(0)

        
def petition_review():
    now = datetime.datetime.now()
    date = now.strftime("%B") + " " + str(now.day) + ", " + str(now.year)
    mhp = validate.string("Enter the name of the mobile home park.").title()
    # First requirement, submitted within 30 days of last meeting
    prompt = ("Was the petition submitted within 30 days of the last "
              "scheduled meeting\nwith the park owner pursuant to FS "
              "723.037(5)(a)?")
    r1 = validate.boolean(prompt)
    if r1 is False:
        print "Too bad. This is a dealbreaker and the petition must be denied."
        letter_center("denied", mhp, "", "")
        exit(0)
    # Second requirement, majority of affected homeowners signed
    lots = validate.integer("Enter the number of affected homeowners.")
    sigsNeeded = lots / 2 + 1
    prompt = ("The homeowners need " + str(sigsNeeded) + " lot designations "
              "to qualify for mediation. Begin\nentering lot numbers now. Do "
              "not count any signatures that do not include\na lot number. "
              "\n\nTo delete the last entry, enter \"del\"."
              "\nWhen you are finished, enter \"end\".\n")
    print "\n" + prompt
    designations = count_designations(lots, sigsNeeded)
    uniques = designations.pop(0)
    duplicates = designations.pop(0)
    print "\nTotal unique signatures: %s (%s needed to approve)" \
           % (str(uniques), str(sigsNeeded))
    r2 = True if uniques >= sigsNeeded else False
    # Third requirement, homeowners indicate reason for petitioning
    prompt = ("1. Rental increase is unreasonable. \n"
              "2. Rental increase has made the lot rental amount"
              " unreasonable.\n"
              "3. Decrease in services or utilities is not accompanied by a\n"
              "   corresponding decrease in rent or is otherwise unreasonable."
              "\n4. Change in rules and regulations is unreasonable.\n\n"
              "Have the majority of the affected homeowners designated in\n"
              "writing one or more of the above reasons to petition for\n"
              "mediation?")
    r3 = validate.boolean(prompt)
    # Fourth requirement, homeowners submit copy of notice being challenged
    prompt = ("Did the homeowners submit a copy of the notice(s) being\n"
              "challenged which indicates a lot rental increase, reduction\n"
              "in services, and/or change(s) to the rules and regulations?")
    r4 = validate.boolean(prompt)
    # Fifth requirement, verification of committee selection form submitted
    prompt = ("Did the homeowners submit a copy of the records verifying the\n"
              "selection of the homeowners' committee in accordance with\n"
              "Florida Statute 723.037(4) and Florida Administrative Rule\n"
              "61B-32.003?")
    r5 = validate.boolean(prompt)
    mediation_review_report(date, mhp, r1, r2, r3, r4, r5, designations,
                            uniques, duplicates, lots, sigsNeeded)
    if r2 is True and r3 is True and r4 is True and r5 is True:
        medInfo = select_mediator()
        letter_center("approved", mhp, medInfo, "", date)
    elif r2 is False or r3 is False or r4 is False or r5 is False:
        reviewPackage = [False, False, False, False, 0, 0, 0]
        reviewPackage[0] = r2
        reviewPackage[1] = r3
        reviewPackage[2] = r4
        reviewPackage[3] = r5
        reviewPackage[4] = lots
        reviewPackage[5] = uniques
        reviewPackage[6] = duplicates
        letter_center("deficient", mhp, "", reviewPackage)


def count_designations(lots, sigsNeeded):
    lotCounter = []
    uniques = 0
    duplicates = 0
    uniqueOrDupe = []
    x = 0
    while x < lots:
        lot = raw_input(str(x + 1).rjust(3) + ". ")
        while lot == "" or lot == " ":
            print "You need to enter something. "
            lot = raw_input(str(x + 1).rjust(3) + ". ")
        if lot == "end":
            break
        elif lot == "delete" or lot == "del":
            if uniqueOrDupe[-1] == "unique":
                uniques -= 1
                del uniqueOrDupe[-1]
            elif uniqueOrDupe[-1] == "duplicate":
                duplicates -= 1
                del uniqueOrDupe[-1]
            del lotCounter[-1]
            x -= 1
            continue
        elif lot in lotCounter:
            uniqueOrDupe.append("duplicate")
            duplicates += 1
            print "Duplicate! There",
            print "is" if duplicates == 1 else "are",
            print "currently " + str(duplicates) + "."
        elif lot not in lotCounter:
            uniqueOrDupe.append("unique")
            uniques += 1
        lotCounter.append(lot)
        x += 1
    lotCounter.insert(0, uniques)
    lotCounter.insert(1, duplicates)
    return lotCounter


def mediation_review_report(date, mhp, r1, r2, r3, r4, r5, designations,
                            uniques, duplicates, lots, sigsNeeded):
    filename = mhp.replace(" ", "") + "PetitionReview"
    scriptDir = os.path.dirname(__file__)
    relPath = "output_files/" + filename + ".txt"
    reportFilePath = os.path.join(scriptDir, relPath)
    report = open(reportFilePath, "w")
    report.write("               MEDIATION REVIEW REPORT\n")
    report.write("\nPetition submitted by homeowners of: %s" % mhp)
    report.write("\n     Mediation petition reviewed by: %s" % NAME)
    report.write("\n     Date petition review completed: %s" % date)
    report.write("\n")
    report.write("\n  Petition submitted within 30 days: ")
    report.write("Yes" if r1 is True else "No")
    report.write("\n          Reason for petition noted: ")
    report.write("Yes" if r3 is True else "No")
    report.write("\n            Copy of notice included: ")
    report.write("Yes" if r4 is True else "No")
    report.write("\n   Committee selection verification: ")
    report.write("Yes" if r5 is True else "No")
    report.write("\n")
    report.write("\n         Number of lots in the park: %s" % lots)
    report.write("\n  Number of lot designations needed: %s" % sigsNeeded)
    report.write("\n   Number of valid lot designations: %s" % uniques)
    report.write("\n\n")
    report.write("The following lot designations were noted in the petition:\n")
    for lot in designations:
        report.write(lot + "\n")
    report.write("\n     Unique designations: %s" % uniques)
    report.write("\n  Duplicate designations: %s\n\n" % duplicates)
    report.write("Duplicate lot designations were not counted for mediation.")
    report.close()


def versa_entry_assistant(date, mhp, circuit, petitioner, petAddress,
                          petCityStateZip, mediator, parkAddress):  
    scriptDir = os.path.dirname(__file__)
    # Get mediator license number
    relPath = "data_files/mediator_license_numbers.txt"
    mediator = HumanName(mediator)
    absFilePath = os.path.join(scriptDir, relPath)
    license = open(absFilePath, "r")
    licenseNumbers = license.read().split("\n")
    for line in licenseNumbers:
        if mediator.last.lower() in line:
            allText = string.maketrans('', '')
            allDigs = allText.translate(allText, digits)
            licenseNumber = line.translate(allText, allDigs)
            break
        else:
            licenseNumber = "Not found!"

    # Output data prep
    filename = mhp.replace(" ", "") + "VersaEntryData"    
    relPath = "output_files/" + filename + ".txt"
    dataFilePath = os.path.join(scriptDir, relPath)
    versa = open(dataFilePath, "w")

    # The document itself
    versa.write(mhp + " Relevant Data for Versa Entry\n\n")
    versa.write("This document will help you enter your data into Versa.\n\n")
    versa.write("First use this info to create an 8304 license:\n\n")
    versa.write("\t" + petitioner + "\n")
    versa.write("\t" + petAddress + "\n")
    versa.write("\t" + petCityStateZip + "\n\n")
    versa.write("Then you'll find in Versa the following license numbers:\n\n")
    versa.write("Respondent (Lic. Type 8301): \n")
    versa.write("Petitioner (Lic. Type 8304): \n")
    versa.write("Park (Lic. Type 8102: \n")
    versa.write("Mediator (Lic. Type 8103): \n\n")
    versa.write("Next you'll need this information to enter the petition:\n\n")
    versa.write(mhp + "\n" + parkAddress + "\n\n")
    versa.write("\tJudicial circuit: " + circuit + "\n")
    versa.write("\tMediator: " + str(mediator) + "\n")
    versa.write("\tMediator License Number: %s" % licenseNumber)
    versa.write("\tApproval Date: %s" % date)
    versa.close()


# Mediatior functions start here # # # # # # # # # # # # # # # # # # # # # # #


def select_mediator():
    result = determine_judicial_circuit()
    city = result.translate(None, digits).title().strip()
    circuit = result[-3:].strip()
    scriptDir = os.path.dirname(__file__)
    relPath = "data_files/mediators.txt"
    absFilePath = os.path.join(scriptDir, relPath)
    mediators = open(absFilePath, "r")
    linesRead = 0 # Counts how many read *before* selected mediator
    while True:
        line = mediators.readline()
        linesRead += 1
        if line[0] != "[":
            continue
        elif line[0] == "[":
            medCircuits = line.strip()
            line = line.split(" ")
            if circuit in line:
                line = mediators.readline()
                mediator = line.strip()
                line = mediators.readline()
                medAddress = line.strip()
                line = mediators.readline()
                medCityStateZip = line.strip()
                line = mediators.readline()
                medPhone = line.strip()
                line = mediators.readline()
                medEmail = line.strip()
                linesRead -= 1
                break
    print "\nHere is the information for the selected mediator:\n"
    print "%s\n%s\n%s\n%s\n%s\n%s\n" % (medCircuits, mediator, medAddress,
                                          medCityStateZip, medPhone, medEmail)
    mediators.close()
    selected = update_mediator_list(mediator, medCircuits, medAddress,
                         medCityStateZip, medPhone, medEmail, linesRead)
    cor.mediator_email(mediator, medAddress, medCityStateZip, medPhone,
                       medEmail, circuit, city, NAME, TITLE, PHONE, EMAIL)
    selected += circuit
    return selected


def update_mediator_list(mediator, medCircuits, medAddress, medCityStateZip,
                         medPhone, medEmail, linesRead):
    scriptDir = os.path.dirname(__file__)
    relPath = "data_files/mediators.txt"
    originalFilePath = os.path.join(scriptDir, relPath)
    fileLength = validate.file_len(originalFilePath)
    mediators = open(originalFilePath, "r")
    relPath = "data_files/updatedMediators.txt"
    updateFilePath = os.path.join(scriptDir, relPath)
    update = open(updateFilePath, "w")    
    for _ in range(0, linesRead):
        line = mediators.readline()
        update.write(line)
    for _ in range(linesRead, linesRead + 7): # skips selected mediator
        line = mediators.readline() 
    for _ in range(linesRead + 6, fileLength + 1): # not sure why 2 and not 1
        line = mediators.readline()
        update.write(line)
    selected = ("\n" + medCircuits + "\n" +
                mediator + "\n" +
                medAddress + "\n" +
                medCityStateZip + "\n" +
                medPhone + "\n" +
                medEmail + "\n")
    update.write(selected)
    mediators.close()
    update.close()
    os.remove(originalFilePath)
    os.rename(updateFilePath, originalFilePath)
    return selected

    
def determine_judicial_circuit():
    city = validate.string("Which city is the park located in?").lower()
    scriptDir = os.path.dirname(__file__)
    relPath = "data_files/circuits.txt"
    circuitFilePath = os.path.join(scriptDir, relPath)
    circuits = open(circuitFilePath, "r")
    while True:
        line = circuits.readline()
        if city in line:
            text = ("\n" + line.translate(None, digits).title().strip() + 
                    " is in judicial circuit " + line[-3:].strip() + ".")
            print text
            circuits.close()
            return line    


# Document generation functions start here # # # # # # # # # # # # # # # # # # 


def letter_center(status, mhp, mediatorInfo, reviewPackage, date=""):
    if status == "":
        prompt = ("Which letter would you like to generate?\n"
                  "1. Denial\n"
                  "2. Deficient\n"
                  "3. Appointment\n")
        choice = validate.menu(prompt, "123")
        if choice == "1":
            status = "denied"
        elif choice == "2":
            status = "deficient"
        elif choice == "3":
            status = "approved"
    if mhp == "":
        mhp = raw_input("\nEnter the name of the mobile home park.\n> ").title()
    if mediatorInfo == "" and status == "approved":
        medFound = validate.boolean("Has a mediator been selected?")
        if medFound is True:
            mediator = validate.string("Enter the mediator's full name, "
                                       "including honorary.").title()
            medAddress = validate.string("Enter the mediator's "
                                         "address.").title()
            medCityStateZip = validate.string("Enter the mediator's city, "
                                              "street, and zip code.").title()
            medPhone = validate.string("Enter the mediator's phone number.")
            medEmail = validate.string("Enter the mediator's email address.")
            circuit = validate.string("Enter the judicial circuit number.")
        if medFound is False:
            mediatorInfo = select_mediator()
            broken = mediatorInfo.split("\n")
            mediator = broken[2]
            medAddress = broken[3]
            medCityStateZip = broken[4]
            medPhone = broken[5]
            medEmail = broken[6]
            circuit = broken[7]
    elif mediatorInfo != "" and status == "approved":
        broken = mediatorInfo.split("\n")
        mediator = broken[2]
        medAddress = broken[3]
        medCityStateZip = broken[4]
        medPhone = broken[5]
        medEmail = broken[6]
        circuit = broken[7]
    petitioner = raw_input("\nEnter the petitioner's full name, including "
                           "honorary.\n> ").title()
    petAddress = raw_input("Enter the petitioner's street address.\n> ").title()
    petCityStateZip = raw_input("Enter the petitioner's city, state, and zip "
                                "code.\n> ").title()
    if status == "denied":
        lastMeeting = raw_input("What was the date of the last meeting?"
                                "\n> ").title()
        submissionDate = raw_input("What date was the petition submitted?"
                                   "\n> ").title()
        cor.denial_letter(mhp, petitioner, petAddress, petCityStateZip,
                          lastMeeting, submissionDate)
    if status == "approved":
        parkAddress = validate.string("Enter the park's mailing address.")
        parkAddress += "\n"
        parkAddress += validate.zip_find("park")
        petContact = validate.string("Enter the petitioner's phone number "
                                     "and/or email address.\n> ")
        medNum = validate.integer("Enter the mediation number.")
        versa_entry_assistant(date, mhp, circuit, petitioner, petAddress,
                              petCityStateZip, mediator, parkAddress)
        cor.appointment_letter(mediator, medAddress, medCityStateZip, medEmail,
                               medPhone, mhp, medNum, petitioner, petContact)
    if status == "deficient" and reviewPackage == "":
        reqs = [True, True, True, True]
        done = False
        while done is False:
            prompt = ("\nSelect a deficiency.\n\n"
                      "1. Didn't submit enough lot designations.\n"
                      "2. Didn't note reason for petitioning.\n"
                      "3. Didn't submit a copy of notice being challenged.\n"
                      "4. Didn't submit committee selection documentation.\n"
                      "5. There are no more deficiencies to enter.")
            answer = validate.menu(prompt, "12345")
            if reqs[0] == False and reqs[1] == False and \
               reqs[2] == False and reqs[3] == False:
                done = True
                break
            elif answer == "1":
                reqs[0] = False
                lots = validate.integer("How many affected homeowners?")
                sigsNeeded = lots / 2 + 1
                uniques = validate.integer("How many valid signatures?")
                duplicates = validate.integer("How many duplicate signatures?")
            elif answer == "2":
                reqs[1] = False
            elif answer == "3":
                reqs[2] = False
            elif answer == "4":
                reqs[3] = False
            elif answer == "5":
                done = True
        if reqs[0] is True:
            lots = 0
            sigsNeeded = 0
            uniques = 0
            duplicates = 0
        cor.deficiency_letter(petitioner, petAddress, petCityStateZip, mhp,
                              lots, sigsNeeded, uniques, duplicates, reqs)
    elif status == "deficient" and reviewPackage != "":
        reqs = [False, False, False, False]
        reqs[0] = reviewPackage[0]
        reqs[1] = reviewPackage[1]
        reqs[2] = reviewPackage[2]
        reqs[3] = reviewPackage[3]
        lots = reviewPackage[4]
        uniques = reviewPackage[5]
        duplicates = reviewPackage[6]
        sigsNeeded = lots / 2 + 1
        cor.deficiency_letter(petitioner, petAddress, petCityStateZip, mhp,
                              lots, sigsNeeded, uniques, duplicates, reqs)

        
if __name__== "__main__":
    main()
