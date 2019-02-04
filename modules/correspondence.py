import docx, os
import win32com.client as win32
from nameparser import HumanName
from mailmerge import MailMerge
from datetime import date

def mediator_email(mediator, medAddress, medCityStateZip, medPhone, medEmail,
                   circuit, city, NAME, TITLE, PHONE, EMAIL):
    mediator = HumanName(mediator)
    statutes = ("http://www.leg.state.fl.us/statutes/index.cfm?App_mode="
                "Display_Statute&URL=0700-0799/0723/0723.html")
    rules = "https://www.flrules.org/gateway/ChapterHome.asp?Chapter=61B-32"
    subject = "Mediation Opportunity in %s" % city
    text =  ("<body style = \"font-family: Calibri;\">"
             "Dear " + mediator.title + " " + mediator.last + ":<br/><br/>"
             "A petition for mediation for a mobile home park in " + city +
             " has been submitted to and approved by the Division of "
             "Condominiums, Timeshares, and Mobile Homes. We are seeking a "
             "mediator to handle this matter. Would you be interested in and "
             "available to mediate a mobile home dispute pursuant to the "
             "applicable <a href=\"" + statutes + "\">Florida Statutes</a> "
             "(sections 723.037 and 723.038) and "
             "<a href=\"" + rules + "\">Florida Administrative Rules</a>?"
             "<br/><br/>Please let me know either way.<br/><br/>"
             "Thanks,<br/>"
             "<strong>" + NAME + "</strong><br/>"
             "<strong>" + TITLE + "</strong><br/>"
             "Department of Business and Professional Regulation<br/>"
             "Division of Florida Condominiums, Timeshares, and Mobile Homes"
             "<br/>Bureau of Compliance<br/>"
             "Phone: " + PHONE + "<br/>"
             "Email: " + EMAIL + "</body>")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = medEmail
    mail.Subject = subject
    mail.HTMLBody = text
    mail.Display(True)

def denial_letter(mhp, petitioner, petAddress, petCityStateZip, lastMeeting,
                  submissionDate):
    name = HumanName(petitioner)
    titleLast = name.title + " " + name.last
    scriptDir = os.path.dirname("mediation.py")
    relPath = "letter_templates/DenialLetterTemplate.docx"
    denialFilePath = os.path.join(scriptDir, relPath)
    document = MailMerge(denialFilePath)
    document.merge(
        mhp = mhp,
        petitioner = petitioner,
        petAddress = petAddress,
        petCityStateZip = petCityStateZip,
        titleLast = titleLast,
        lastMeeting = lastMeeting,
        submissionDate = submissionDate)
    docName = mhp.replace(" ", "") + "DenialLetter.docx"
    relPath = "output_files/" + docName
    outputFilePath = os.path.join(scriptDir, relPath)
    document.write(outputFilePath)
    create_pdf(outputFilePath) 
    print "\nA denial letter has been created!\n"   

def deficiency_letter(petitioner, petAddress, petCityStateZip, mhp, lots,
                      sigsNeeded, uniques, duplicates, reqs):
    scriptDir = os.path.dirname("mediation.py")
    count = 0
    for deficiency in reqs:
        if deficiency is True:
            count += 1
    name = HumanName(petitioner)
    titleLast = name.title + " " + name.last
    sigsShort = sigsNeeded - uniques
    r2, r3, r4, r5 = "", "", "", "" # equivalent to reqs[0] through reqs[3]
    if reqs[0] is False:
        relPath = "letter_templates/Requirement2Template.txt"
        r2FilePath = os.path.join(scriptDir, relPath)
        data = open(r2FilePath)
        r2 = data.read()
        r2 = "\n\n" + r2 % (str(lots), str(sigsNeeded), str(uniques),
                   str(duplicates), str(sigsShort))
        data.close()
    if reqs[1] is False:
        relPath = "letter_templates/Requirement3Template.txt"
        r3FilePath = os.path.join(scriptDir, relPath)
        data = open(r3FilePath)
        r3 = "\n\n" + data.read()
        data.close()
    if reqs[2] is False:
        relPath = "letter_templates/Requirement4Template.txt"
        r4FilePath = os.path.join(scriptDir, relPath)
        data = open(r4FilePath)
        r4 = "\n\n" + data.read()
        data.close()
    if reqs[3] is False:
        relPath = "letter_templates/Requirement5Template.txt"
        r5FilePath = os.path.join(scriptDir, relPath)
        data = open(r5FilePath)
        r5 = "\n\n" + data.read()
        data.close()
    singleOrPlural = "deficiency" if count == 1 else "deficiencies"
    relPath = "letter_templates/DeficiencyLetterTemplate.docx"
    deficientFilePath = os.path.join(scriptDir, relPath)
    document = MailMerge(deficientFilePath)
    document.merge(
        petitioner = petitioner,
        petAddress = petAddress,
        petCityStateZip = petCityStateZip,
        mhp = mhp,
        titleLast = titleLast,
        singleOrPlural = singleOrPlural,
        r2 = r2,
        r3 = r3,
        r4 = r4,
        r5 = r5)
    docName = mhp.replace(" ", "") + "DeficiencyLetter.docx"
    relPath = "output_files/" + docName
    outputFilePath = os.path.join(scriptDir, relPath)
    document.write(outputFilePath)
    create_pdf(outputFilePath) 
    print "\nA deficiency letter has been created!\n"
        
def appointment_letter(mediator, medAddress, medCityStateZip, medEmail,
                       medPhone, mhp, medNum, petitioner, petContact):
    name = HumanName(mediator)
    titleLast = name.title + " " + name.last
    scriptDir = os.path.dirname("mediation.py")
    relPath = "letter_templates/AppointmentLetterTemplate.docx"
    appointmentFilePath = os.path.join(scriptDir, relPath)
    document = MailMerge(appointmentFilePath)
    document.merge(
        mediator = mediator,
        medAddress = medAddress if medAddress != "N/A" else "",
        medCityStateZip = medCityStateZip if medCityStateZip != "N/A" else "",
        medEmail = medEmail,
        medPhone = medPhone,
        mhp = mhp,
        medNum = str(medNum),
        titleLast = titleLast,
        petitioner = petitioner,
        petContact = petContact)
    docName = mhp.replace(" ", "") + "AppointmentLetter" + str(medNum) + ".docx"
    relPath = "output_files/" + docName
    outputFilePath = os.path.join(scriptDir, relPath)
    document.write(outputFilePath)
    create_pdf(outputFilePath) 
    print "\nAn appointment letter has been created!\n"    

def create_pdf(outputFilePath):
    print "Please wait just one moment... "
    wdFormatPDF = 17
    in_file = os.path.abspath(outputFilePath)
    pdfFilePath = outputFilePath.replace(".docx", ".pdf")
    out_file = os.path.abspath(pdfFilePath)
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
