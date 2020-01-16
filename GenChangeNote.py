import sys
import requests
from requests.auth import HTTPDigestAuth
import json
import datetime
import csv
import xlsxwriter

GerritURL="MDT-GERRIT01.mic.com.tw:8888"
GerritAuth="aken.hsu:AAaa7410"
DefaultFileName="ReleaseNote"

ItmBrch="Branch"
ItmRvNo="ReviewNo"
ItmProj="Project"
ItmSubj="Subject"
ItmSmtr="Owner"
ItmSmtD="SubmitDate"
ChangeItems=[ItmBrch, ItmRvNo, ItmProj, ItmSubj, ItmSmtr, ItmSmtD]
ItmBgFx="BugFix"
ItmFthr="Feature"
MessageItems=[ItmBgFx, ItmFthr]

def parseCommitMessage(CommitMessage, Change):
    print(CommitMessage)
    Items = CommitMessage.split("\n")
    print(Items)

def getCommitDetail(RevNo, Change):
    RequestStr = "http://%s@%s/a/changes/%s/detail" % (GerritAuth, GerritURL, RevNo)
    print(RequestStr)
    sys.stdout.write("Accessing %s ... " % RequestStr)
    Resp = requests.get(RequestStr)
    if (Resp.ok):
        sys.stdout.write("Request OK.\n")
        RespCont = Resp.content.split("\n",1)[1]; # Remove 1st line with the ")]}'" unnecessary chars and will make json not formatted
        jData = json.loads(RespCont)
        # print RespCont
        Change.append(jData["branch"])
        Change.append(RevNo)
        Change.append(jData["project"])
        Change.append(jData["subject"])
        Change.append(jData["owner"]["email"])
        Change.append(jData["submitted"].split(".")[0])
    else:
        sys.stdout.write("Request NG!\n")

    # Find the lasest commit message
    for i in range(1, 99):
        RequestStr = "http://%s@%s/a/changes/%s/revisions/%s/commit" % (GerritAuth, GerritURL, RevNo, i)
        sys.stdout.write("Accessing %s ... " % RequestStr)
        Resp = requests.get(RequestStr)
        if (Resp.ok):
            sys.stdout.write("Request OK.\n")
            RespCont = Resp.content.split("\n",1)[1]; # Remove 1st line with the ")]}'" unnecessary chars and will make json not formatted
            # print RespCont
        else:
            sys.stdout.write("Request NG!\n")
            break

    jData = json.loads(RespCont)
    CommitMessage = jData["message"]
    parseCommitMessage(CommitMessage, Change)

def writeAsCsvFile(ReleaseNote, FileName):
    with open(FileName, 'w') as fp:
        Writer = csv.writer(fp)
        Writer.writerow(ChangeItems)
        for EachC in ReleaseNote:
            Writer.writerow(EachC)

def writeAsExcelFile(ReleaseNote, QueryConditions, FileName):
    ExcelFile = xlsxwriter.Workbook(FileName)

    # Write change data as a sheet
    ChangeSheet = ExcelFile.add_worksheet('Changes')
    Row = 0
    Colum = 0
    for EachI in ChangeItems: # Write the first row that descripts what the colume is 
        ChangeSheet.write(Row, Colum, EachI)
        Colum += 1
        
    for EachC in ReleaseNote: # Write the Changes for each Row
        Row += 1
        Colum = 0
        for EachI in EachC:
            ChangeSheet.write(Row, Colum, EachI)
            Colum += 1

    # Write query data as a sheet
    QuerySheet = ExcelFile.add_worksheet('Query')
    Row = 0
    Colum = 0
    for EachQ in QueryConditions:
        print("%s - %s" % (EachQ, QueryConditions[EachQ]))
        QuerySheet.write(Row, 0, EachQ)
        QuerySheet.write(Row, 1, QueryConditions[EachQ])
        Row += 1
        
    ExcelFile.close()

def writeReleaseNote(ReleaseNote, QueryConditions, FileName):
    print(ReleaseNote)
    CsvFileName="%s.csv" % FileName
    writeAsCsvFile(ReleaseNote, CsvFileName)
    ExcelFileName="%s.xlsx" % FileName
    writeAsExcelFile(ReleaseNote, QueryConditions, ExcelFileName)

def handleQueryChange(QueryStr, ReviewNOs):
    RequestStr="http://%s@%s/a/changes/?q=%s" % (GerritAuth, GerritURL, QueryStr)
    sys.stdout.write("Accessing %s ... " % RequestStr)
    Resp=requests.get(RequestStr)
    if (Resp.ok):
        sys.stdout.write("Request OK.\n")
        RespCont=Resp.content.split("\n",1)[1]; # Remove 1st line with the ")]}'" unnecessary chars and will make json not formatted
        jData=json.loads(RespCont)
        for EachR in jData:
            ReviewNOs.append(EachR["_number"])

        return 0

    else:
        sys.stdout.write("Request NG!\n")
        return 1

def handleReviewNOs(ReviewNOs):
    ReleaseNote = [] # A container for many Changes
    for EachN in ReviewNOs:
        Change=[] # A container for many information about a Change
        getCommitDetail(EachN, Change)
        ReleaseNote.append(Change)
    return ReleaseNote

def parseQueryMessage(QueryMessage):
    QueryConditions = {}
    for EachS in QueryMessage.split("+"):
        EachQ = EachS.split(":", 1)
        QueryConditions[EachQ[0]] = EachQ[1]
    return QueryConditions

DefaultBranch="sc20-android-quectel-evb"
DefaultBegin="2019-11-12 0:0:0"
DefaultFinish="2019-11-20 0:0:0"
def main():
    try: Branch = sys.argv[1]
    except IndexError: Branch = DefaultBranch
    try: Begin = sys.argv[2]
    except IndexError: Begin = DefaultBegin
    try: Finish = sys.argv[3]
    except IndexError: Finish = DefaultFinish
    try: FileName = sys.argv[4]
    except IndexError: FileName = DefaultFileName    

    QueryStr = "branch:%s+after:%s+before:%s+status:merged" %(Branch, Begin, Finish)
    print(QueryStr)

    if len(sys.argv) < 4:
        sys.stderr.write('Usage: %s Branch "BegineTime" "FinishTime" like:\n%s %s "%s" "%s"\n' % (sys.argv[0], sys.argv[0], Branch, Begin, Finish))
        return 1

    ReviewNOs=[] # A container for ChangeIDs
    ret = handleQueryChange(sys.argv[1], ReviewNOs)
    if (ret != 0): return 1
    print(ReviewNOs)
    
    ReleaseNote = handleReviewNOs(ReviewNOs)
    print(ReleaseNote)

    QueryConditions = parseQueryMessage(QueryStr)
    print(QueryConditions)

    writeReleaseNote(ReleaseNote, QueryConditions, FileName)

    return 0

if __name__ == '__main__':
    sys.exit(main())
