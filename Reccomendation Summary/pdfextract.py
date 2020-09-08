import sys, os
import pyexcel
from pathlib import Path
import fitz

#print("---------------------------------------------------------------------------------------------------")
content = []

def split_last_element(stringList, delim, maxSplit = 1):
    length = len(stringList)
    tempSplitList = stringList[length - 1].split(delim, maxSplit)
    stringList.pop()
    stringList.extend(tempSplitList)

def remove_left(string, delim):
        if delim == "":
            return string
        else:
            return string.split(delim, 1)[1]

def strip_ws(string):
    return string.replace(" ", "").replace("/n","")

def debugPrint(*varsToPrint, arr=content):
    i = 0
    for var in varsToPrint:
        i += 1
        print(str(i) + ") " + var.strip())
    line = ""
    for _ in range(38):
        line += " --"
    print(line)
    print(arr[0] if len(arr) else "CONTENT EMPTY")
    print(line)

def extract(stringList, signPostAfter, signPostBefore="", txtToReplace = "", txtToInsert = "", i = 0):
    if (signPostAfter not in stringList[i] or signPostBefore not in stringList[i]):
        raise IndexError

    stringList[i] = remove_left(stringList[i], signPostBefore)
    stringList[i] = stringList[i].replace(txtToReplace,txtToInsert,1)
    toReturn, stringList[i] = stringList[i].split(signPostAfter,1)[0], stringList[i].split(signPostAfter,1)[1]
    return toReturn

def rip_pdf_text(filePath):
    #Rip PDF
    #processed_file = parser.from_file(filePath, 'http://localhost:3000/tika')
    #content = processed_file['content']
    doc = fitz.open(filePath)
    content = ""
    for page in doc:
        content += page.getText()

    #Chunk Data
    content = "".join([s.replace("\n"," ") for s in content.splitlines(True) if s.strip()]) #removes empty lines
    content = content.split(    "Missionary Recommendation"                                , 3)
    split_last_element(content, "Education and Service of Missionary Candidate"               )
    split_last_element(content, "Unit Information for Missionary Candidate"                   )
    split_last_element(content, "Priesthood Leaders' Comments and Suggestions"                )
    split_last_element(content, "Personal Health History of Missionary Candidate"             )
    #if ("Missionary Additional Health History" in content):
    #    split_last_element(content, "Missionary Additional Health History"                        )
    split_last_element(content, "Physician's Health Evaluation for Prospective Missionary"    )
    split_last_element(content, "Personal Insurance Information of Missionary Candidate"      )
    content.pop(0)
    content[1] = strip_ws(content[1])
    return content

def dataPull(arrayStrings):
    #extract(stringList, signPostAfter, signPostBefore, txtToReplace, txtToInsert, i)
    #PAGE 1 - Missionary Reommendation
    homeStreetAddr = extract(arrayStrings, "City","Home street address").strip()
    homeCity = extract(arrayStrings,"State or province").strip()
    homeState = extract(arrayStrings,"Postal code").strip()
    homeZip = extract(arrayStrings," Country").strip()
    homePhone = extract(arrayStrings, " Mobile phone", "area code)")
    DOB = extract(arrayStrings, "Gender", "Date of birth")
    arrayStrings.pop(0)
    
    #PAGE 2 - Missionary Reommendation
    try:
        givenNames = extract(arrayStrings, "(last)", "Yourfulllegalname(first)", "(middle)", " ")
    except IndexError as newForm:
        givenNames = extract(arrayStrings, "LastName", "FirstName", "(middle)", " ")
    
    surname = extract(arrayStrings, "(suffix)")
    memRecordNum = extract(arrayStrings, "Dateofbirth", "Recordnumber")
    arrayStrings.pop(0)

    #PAGE 3 - Missionary Reommendation
    #father
    try:
        fatherName = extract(arrayStrings, "Father is a member", "Father's full name")
    except IndexError as newForm:
        fatherName = extract(arrayStrings, "Middle", "First Name") + extract(arrayStrings, "Last Name") + extract(arrayStrings, "Father is a member")

    fatherJob = extract(arrayStrings, "Father's street", "Father's occupation")
    fatherAddr = "" if not arrayStrings[0].find("City State or province", 0, 68) + 1 \
                    else ( extract(arrayStrings, "City","home address").strip() + 
                        "\n" +  extract(arrayStrings,"State or province").strip()  +  ", "  +
                        extract(arrayStrings,"Postal code").strip() + " " + extract(arrayStrings," Country").strip() )
    fatherPhone = extract(arrayStrings,"Check here if you do NOT want your father" ,"Mobile phone (indicate country and include area code)")

    #mother
    try:
        motherName = extract(arrayStrings, "Mother is a member", "Mother's full name")
    except IndexError as newForm:
        motherName = extract(arrayStrings, "Middle", "First Name") + extract(arrayStrings, "Last Name") + extract(arrayStrings, "Mother is a member")

    motherJob = extract(arrayStrings, "Mother's street", "Mother's occupation")
    motherAddr = "" if not arrayStrings[0].find("City State or province", 0, 68) + 1 \
                    else ( extract(arrayStrings, "City","home address").strip() + 
                        "\n" +  extract(arrayStrings,"State or province").strip()  +  ", "  + 
                        extract(arrayStrings,"Postal code").strip() + " " + extract(arrayStrings," Country").strip() )
    motherPhone = extract(arrayStrings,"Check here if you do NOT want your mother" ,"Mobile phone (indicate country and include area code)")

    #other
    try:
        otherName = extract(arrayStrings, "(relationship)", "Guardian (Other)").strip()
    except IndexError as newForm:
        otherName = extract(arrayStrings, "Relationship", "Guardian (Other)").strip()

    if otherName:
        otherAddr = "" if not arrayStrings[0].find("City State or province", 0, 68) + 1 \
                    else ( extract(arrayStrings, "City","home address").strip() + 
                        "\n" +  extract(arrayStrings,"State or province").strip()  +  ", "  + 
                        extract(arrayStrings,"Postal code").strip() + " " + extract(arrayStrings," Country").strip() )
        otherPhone = extract(arrayStrings,"Check here if you do NOT want this person" ,"Mobile phone (indicate country and include area code)")
    else:
        otherAddr = ""
        otherPhone = ""

    residesWith = ""
    if not fatherAddr.strip() and not motherAddr.strip(): #mom and dad both blank
        residesWith = "Parents"
        fatherAddr = motherAddr = homeStreetAddr + " " + homeCity + ", " + homeState + " " + homeZip
    elif not fatherAddr.strip(): #dad blank
        residesWith = "Father"
        fatherAddr = homeStreetAddr + " " + homeCity + ", " + homeState + " " + homeZip
    elif not motherAddr.strip(): #mother blank
        residesWith = "Mother"
        motherAddr = homeStreetAddr + " " + homeCity + ", " + homeState + " " + homeZip
    else:
        residesWith = "Other"
        otherAddr = homeStreetAddr + " " + homeCity + ", " + homeState + " " + homeZip
        
    #missions
    fatherMission = extract(arrayStrings, "Mother has served a mission. Yes No", "Father has served a mission Yes No")
    motherMission = extract(arrayStrings, "Grandparents have served missions Yes No")
    grandparentsMission = extract(arrayStrings, "Do you have any parent, brother, sister, grandparent, or boyfriend/girlfriend currently serving a mission? If yes, list the name, relationship, and mission for each person.")
    otherMission = extract(arrayStrings, "by Intellectual Reserve, Inc. All rights reserved.") if arrayStrings[0]. find("by Intellectual Reserve") + 1 else arrayStrings[0]
    arrayStrings.pop(0)

    #PAGE 4 - Education and Service of Missionary Candidate
    languages = extract(arrayStrings," Indicate all other languages", "What is your primary language?")
    languages = languages[:len(languages)-(2 if languages.find("-")+1 or languages.find("+")+1 else 1)]
    while(arrayStrings[0].find("Language you want your call letter to be printed in") > 5):
        languages += "; " + extract(arrayStrings, " Language you want your call letter to be printed in","Average grade ")
    education = extract(arrayStrings, "You have earned or will earn:","Highest education level achieved")

    try:
        yrsSeminary = extract(arrayStrings, "Graduated from seminary", "Years of seminary completed")
    except IndexError as newForm:
        yrsSeminary = extract(arrayStrings, "Did you graduate from seminary", "How many years did you attend seminary and/or institute?")
    school1  = [extract(arrayStrings, "Degree", "Number of years"),
                extract(arrayStrings, "School", "Major"),
                extract(arrayStrings, "Number of years")] #yrs,major,schoolname
    school2  = [extract(arrayStrings, "Degree"),
                extract(arrayStrings, "School", "Major"),
                extract(arrayStrings, "Extracurricular activities, special skills, hobbies, and special accomplishments")] #yrs,major,schoolname

    if school1[1].strip() or school1[2].strip() or school2[1].strip() or school2[2].strip():
        education += "| "
        if school1[1].strip() or school1[2].strip():
            education += ", ".join(school1)
            if school2[1].strip() or school2[2].strip():
                education += "; " + ", ".join(school2)
        else:
            if school2[1].strip() or school2[2].strip():
                education += ", ".join(school2)
    
    extraCurr = extract(arrayStrings, "Previous Church callings and leadership experience")
    callings  = extract(arrayStrings, "Work experience outside the home (Include number of years in each job.)")
    try:
        work = extract(arrayStrings, "Office experience General bookkeeping")
    except IndexError as newForm:
        work = extract(arrayStrings, "Office: General bookkeeping")


    #PAGE 5 - Education and Service of Missionary Candidate
    selfFinance = extract(arrayStrings, "Family (per month)", "Self (per month)")
    familyFinance = extract(arrayStrings, "Ward or branch (per month)")
    unitFinance = extract(arrayStrings, "Other (per month)")
    otherFinance = extract(arrayStrings, "Total to be paid per month")
    missionFinance = ""
    if selfFinance.strip():
        missionFinance += "Self - " + selfFinance
    if familyFinance.strip():
        if missionFinance:
            missionFinance += "; "
        missionFinance += "Family - " + familyFinance
    if unitFinance.strip():
        if missionFinance:
            missionFinance += "; "
        missionFinance += "Unit - " + unitFinance
    if otherFinance.strip():
        if missionFinance:
            missionFinance += "; "
        missionFinance += "other - " + otherFinance
    arrayStrings.pop(0)

    #PAGE 6 - Unit Information for Missionary Candidate
    unitLeaderName = extract(arrayStrings, "Name of home stake or mission president", "Name of home bishop or branch president")
    stakePresName = extract(arrayStrings, "Mailing address (including country)")
    unitLeaderAddr = extract(arrayStrings, "Mailing address (including country)").replace("United States", "")
    stakePresAddr = extract(arrayStrings, "Home phone (area code)").replace("United States", "")
    
    unitLeaderHome = extract(arrayStrings, "Work phone (area code)")
    unitLeaderMobile = extract(arrayStrings, "Home phone (area code)", "Cell phone (area code)")
    unitLeaderPhone = unitLeaderMobile if unitLeaderMobile.strip() else unitLeaderHome 
    
    stakePresHome = extract(arrayStrings, "WorkPhoneLabel")
    stakePresMobile = extract(arrayStrings, "E-mail address", "Cell phone (area code)")
    stakePresPhone = stakePresMobile if stakePresMobile.strip() else stakePresHome 
    
    unitLeaderEmail = extract(arrayStrings, "Fax")
    stakePresEmail = extract(arrayStrings, "Fax", "E-mail address")
    arrayStrings.pop(0)

    #PAGE 7 - Priesthood Leaders' Comments and Suggestions
    unitLeaderComments = extract(arrayStrings, "Please evaluate the missionary candidate's leadership capability.", "Confidential comments should be discussed in a separate letter.")

    #PAGE 8 - Priesthood Leaders' Comments and Suggestions
    stakePresComments = extract(arrayStrings, "When you sign this form, you are stating that in your opinion this", "Confidential comments should be discussed in a separate letter.")
    arrayStrings.pop(0)

    #PAGE 9~ - Personal Health History of Missionary Candidate
    arrayStrings.pop(0)

    #PAGE 10 - Missionary Additional Health History
    #arrayStrings.pop(0)

    #PAGE 11 - Physician's Health Evaluation for Prospective Missionary
    height = extract(arrayStrings, "in.", "Height (in inches or centimeters)")
    weight = extract(arrayStrings, "lbs.", "(in pounds or kilograms)")
    arrayStrings.pop(0)

    #PAGE 12 - Personal Insurance Information of Missionary Candidate
    ins_companyName = extract(arrayStrings, "Policyholder's name", "Name of primary insurance company")
    ins_companyName = ins_companyName if ins_companyName.strip() else "No Company Listed"
    if ins_companyName.strip():
        ins_policyHolder = extract(arrayStrings, "Policyholder's date of birth")
        ins_dob = extract(arrayStrings, "Effective date of coverage")
        ins_groupNum = extract(arrayStrings, "Policyholder's ID number", "Policyholder's Group Number")
        ins_policyHolderID = extract(arrayStrings, "Mailing address for submitting claims")
        if not (arrayStrings[0].find("City", arrayStrings[0].find("City", arrayStrings[0].find("City", 0, 500) + 4, 500) + 4, 500) + 1):
            ins_mailingAddr = extract(arrayStrings, "City")
        else:
            secondCityIndex = arrayStrings[0].find("City", arrayStrings[0].find("City", 0, 500) + 4, 500)
            ins_mailingAddr, arrayStrings[0] = arrayStrings[0][0:secondCityIndex], arrayStrings[0][secondCityIndex + 4:]
        ins_city = extract(arrayStrings, "State or province")
        ins_state = extract(arrayStrings, "Postal code")
        ins_zip = extract(arrayStrings, "Country")
        ins_country = extract(arrayStrings, "District (if any) Phone number of insurance company (include area code)")

        ins_phoneNumber = extract(arrayStrings, "Indicate where this insurance plan will")
    else:
        ins_policyHolder = ""
        ins_dob = ""
        ins_groupNum = ""
        ins_policyHolderID = ""
        ins_mailingAddr = ""
        ins_city = ""
        ins_state = ""
        ins_zip = ""
        ins_country = ""
        ins_phoneNumber = ""
    arrayStrings.pop(0)

    varNames = '''Email, DOB, missionaryID, memRecordNum, medicalID, MTCSTART, Arrival, Release, Name, homeAddr, homePhone, age, residesWith, relation, fatherName, fatherAddr, fatherPhone, fatherJob, fatherMission, relation, motherName, motherAddr, motherPhone, motherJob, motherMission, languages, yrsSeminary, education, extraCurr, callings, work, missionFinance, unitLeaderName, unitLeaderAddr, unitLeaderEmail, unitLeaderPhone, unitLeaderComments, stakePresName, stakePresAddr, stakePresEmail, stakePresPhone, stakePresComments, ins_companyName, ins_policyHolder, INS_DOB, ins_groupNum, ins_policyHolderID, ins_addr, ins_city, ins_state, ins_zip, ins_phoneNumber, height, weight, limitations'''.replace(" ", "").upper().split(",")
    compiledData = ["", DOB, "", memRecordNum, "", "", "","", surname.upper() + ", " + givenNames, homeCity + ", " + homeState, homePhone,'\'=DATEDIF(D3,D8,"Y")', residesWith, "Father", 
                    fatherName, fatherAddr, fatherPhone, fatherJob, fatherMission, "Mother", motherName, motherAddr, motherPhone, motherJob, motherMission,
                    languages, yrsSeminary, education, extraCurr, callings, work, missionFinance, unitLeaderName, unitLeaderAddr, 
                    unitLeaderEmail, unitLeaderPhone, unitLeaderComments, stakePresName, stakePresAddr, stakePresEmail, stakePresPhone,
                    stakePresComments, ins_companyName, ins_policyHolder, ins_dob, ins_groupNum, ins_policyHolderID,
                    ins_mailingAddr, ins_city, ins_state, ins_zip, ins_phoneNumber, height, weight, ""]

    for i in range(len(compiledData)):
        compiledData[i] = compiledData[i].strip()

    return varNames, compiledData

def txtExport(varNames, compiledData):
    #TXT EXPORT
    Path(os.path.dirname(os.path.realpath(__file__)) + "/#txtExports/").mkdir(parents=True, exist_ok=True)
    with open(os.path.dirname(os.path.realpath(__file__)) + "/#txtExports/"+ compiledData[8] + ".txt","w") as output:
        i = 0
        for info in compiledData:
            output.write(varNames[i].upper() + f'\n{info}\n\n')
            i += 1       
        #print("check the DNC boxes\nSort out Language"

def exclExport(varNames, compiledData):
    Path(os.path.dirname(os.path.realpath(__file__)) + "/#xlsxExports/").mkdir(parents=True, exist_ok=True)
    sheetData = []
    for i in range(len(compiledData)):
        if i > 18 and i < 25:
            sheetData[i-6].append(compiledData[i])
            sheetData[i-6][0] = sheetData[i-6][0] + " / " + varNames[i]
        else:
            sheetData.append([varNames[i], "", "", compiledData[i]])

    for i in [1, 10, 11, 17, 18, 25, 26, 34, 35, 46, 47, 58, 59]:
        sheetData.insert(i - 1, ["","","",""])

    #print(sheetData)
    pyexcel.save_as(array=sheetData, dest_file_name=os.path.dirname(os.path.realpath(__file__)) + "/#xlsxExports/" + compiledData[8] + ".xlsx")

def main(*args):
    #print(sys.argv)
    if len(sys.argv) > 1:
        for i in range(1, len(sys.argv)):
            currFile = sys.argv[i]
            content = rip_pdf_text(sys.argv[i])
            varNames, compiledData = dataPull(content)

            #txtExport(varNames, compiledData)
            exclExport(varNames, compiledData)
    else:
        input("Exit Now")

if __name__ == '__main__':
    main()