"""
    hgSpreadsheetReporter

     Script for filling mercurial history into Google spreadsheets
"""
__version__ = "0.1"
__author__ = 'Iurii Gazin [archeg]'

import subprocess
import datetime
import re

import dateutil.parser

WeekNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


class ConsoleColors:
    Header = '\033[95m'
    OkBlue = '\033[94m'
    OkGreen = '\033[92m'
    Warning = '\033[93m'
    Fail = '\033[91m'
    EndC = '\033[0m'


def getGitHistory(path):
    """ Returns all the mercurial history, by looking at the path.
    The path should point to the root of mercurial repository
    :param path:
    :return: list of (lRevision, lChangeSet, lUser, lDate, lSummary)
    """
    import os

    os.chdir(path)
    lProcess = subprocess.Popen(["hg","history"], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
    lOutput = lProcess.stdout.readlines()
    lHistory = []
    for row in lOutput:
        row = "".join(row)
        if row.startswith('changeset:'):
            lChangeSet = row[10:].strip()
            lRevision = lChangeSet[:lChangeSet.find(':')].strip()
        elif row.startswith('user:'):
            lUser = row[10:].strip()
        elif row.startswith('date:'):
            lDateText = row[10:].strip()[:-5].strip()
            lDate = datetime.datetime.strptime(lDateText, "%a %b %d %H:%M:%S %Y")
        elif row.startswith('summary:'):
            lSummary = row[10:].strip()
        elif row =='\n':
            lHistory.append((lRevision, lChangeSet, lUser, lDate, lSummary))

    return lHistory


def getConfig():
    """ Reads the config from the configuration file.
    :return: configuration parser.
    """
    import ConfigParser, os
    parser = ConfigParser.RawConfigParser()
    parser.readfp(open("config.ini"))
    return parser

# Sections to be ignored by project searcher. All other sections are considered project names.
ignoredSections = ["GoogleDocs", "General"]

configParser = getConfig()


def readDataAndSplit(configParser, group, key):
    """ Reads data from config parser and splits it by comma
    """
    return [x.strip() for x in configParser.get(group, key).split(",") if x.strip() != ""]


def parseDayStart(s):
    """ Parses a day start acquiring only hours and minutes, and setting date to 2.2.2000
    :param s:
    :return:
    """
    dateMatch = re.match(r"(\d*)[:|\.](\d*)", s)
    return datetime.datetime(2000, month = 2, day = 2, hour = int(dateMatch.group(1)), minute = int(dateMatch.group(2)))


def safeGet(parser, section, option):
    """ Safely gets a value from config parser. If value is not present - returns None
    """
    if parser.has_option(section, option):
        return parser.get(section, option).strip()
    return None


def parseBool(string):
    return string.lower() == "true"

# ------- Reading config here -------------
gdocsUser = configParser.get("GoogleDocs", "user").strip()
gdocsPwd = safeGet(configParser, "GoogleDocs", "pwd")
gdocsName = configParser.get("GoogleDocs", "docName").strip()
gdocsWorksheet = configParser.get("GoogleDocs", "worksheet").strip()
gdocsUserToFill = configParser.get("GoogleDocs", "userToFill").strip()

gdocsTotalKey = configParser.get("GoogleDocs", "totalColumnName").strip()
gdocsNameKey = configParser.get("GoogleDocs", "nameColumnName").strip()
gdocsDateKey = configParser.get("GoogleDocs", "dateColumnName").strip()
gdocsOtherKey = configParser.get("GoogleDocs", "otherColumnName").strip()
gdocsOtherShortcut = configParser.get("GoogleDocs", "otherShortcut").strip()

gdocsCommentsKey = safeGet(configParser, "GoogleDocs", "commentsColumnName")

workdayStart = parseDayStart(configParser.get("General", "workdayStart").strip())
workdayLength = float(configParser.get("General", "workdayLength").strip())
truncateToWorkday = parseBool(configParser.get("General", "truncateToFullworkday").strip())
silentMode = parseBool(configParser.get("General", "silentMode").strip())

projectsData = []
projectShortcutDict = {}

# Reading all sections except ignored, as project setups
for section in configParser.sections():
    if section not in ignoredSections:
        path = configParser.get(section, "path")
        usersToFilter = readDataAndSplit(configParser, section, "users")
        shortcut = configParser.get(section, "projectShortcut")

        projectsData.append((section, path, usersToFilter, shortcut))
        projectShortcutDict[shortcut] = section
projectShortcutDict[gdocsOtherShortcut] = gdocsOtherKey

startDate = datetime.datetime(1988, 11, 7)
endDate = datetime.datetime.now()

# --------------

print "hgSpreadsheetReporter. Version " + str(__version__)
print ""
print "This script will iterate the history of mercurial repositories, noted in configuration, and put the report to " \
      "Google spreadsheet"
print ConsoleColors.Fail + "Please be sure that Google preserve history properly. It is even suggested to make a " \
                           "copy of your Google Spreadsheet. \nThere is no guarantee that this script won't " \
                           "damage your already exist data, as it does pretty aggressive resizing of the spreadsheet." \
                           "\nIt is designed to preserve the old history, but there is no guarantee. \nAlso be sure " \
                           "that Spreadsheet has supported template. Otherwise you probably loose your data. " \
                           "Use this tool on your own risk.!!!" + ConsoleColors.EndC
print ""
print "------------------"
print ""

history = []
for projectName, path, usersToFilter, prjShortcut in projectsData:
    hist = getGitHistory(path)

    hist = [(projectName,  rev, changeSet, date, summary) for rev, changeSet, user, date, summary in hist if user in usersToFilter and startDate < date < endDate]

    history = history.__add__(hist)

# Asking for a password, if not set in config
if gdocsPwd is None or gdocsPwd == "None":
    import getpass
    gdocsPwd = getpass.getpass("Please enter your google password")

# Google spreadsheet
import gspread
gc = gspread.login(gdocsUser, gdocsPwd)
spreadsheet = gc.open(gdocsName)

worksheet = spreadsheet.worksheet(gdocsWorksheet)


# -------- Parsing header ------------

headerRow = worksheet.row_values(1)
#Find all the indexes of columns that we are interested in
nameColumnNo = None
dateColumnNo = None
totalColumnNo = None
commentsColumnNo = None
otherColumnNo = None
projectsColumnNo = {}
for currentNumberOfRows in range(len(headerRow)):
    if headerRow[currentNumberOfRows] == gdocsNameKey:
        nameColumnNo = currentNumberOfRows + 1
    if headerRow[currentNumberOfRows] == gdocsDateKey:
        dateColumnNo = currentNumberOfRows + 1
    if headerRow[currentNumberOfRows] == gdocsTotalKey:
        totalColumnNo = currentNumberOfRows + 1
    if headerRow[currentNumberOfRows] == gdocsCommentsKey:
        commentsColumnNo = currentNumberOfRows + 1
    if headerRow[currentNumberOfRows] == gdocsOtherKey:
        otherColumnNo = currentNumberOfRows + 1
    if headerRow[currentNumberOfRows] in [section for section, path, usersToFilter, prjShortcut in projectsData]:
        projectsColumnNo[headerRow[currentNumberOfRows]] = currentNumberOfRows + 1

# -------- Support methods ----------

def composeADate(date, time):
    """ Composes a date from separate date and time"""
    return datetime.datetime(date.year, date.month, date.day, time.hour, time.minute)


def composeARow(user, date, projectsHours, comments):
    """ Composes a row from given project data to be posted to spreadsheet
    """
    rowLength = max(max(nameColumnNo, dateColumnNo, totalColumnNo, commentsColumnNo, otherColumnNo), max(projectsColumnNo.values()))
    row = ['' for _ in range(rowLength)]

    row[nameColumnNo - 1] = user
    row[dateColumnNo - 1] = date

    # Compose total column value. We will use RC-format for that
    s = []
    for prjColumnNo in projectsColumnNo.values():
        shift = prjColumnNo - totalColumnNo
        s.append("R[0]C[" + str(shift) + "]")
    s.append("R[0]C[" + str(otherColumnNo - totalColumnNo) + "]")

    row[totalColumnNo - 1] = "=" + "+".join(s)

    for project in projectsColumnNo.keys():
        if projectsHours.has_key(project):
            row[projectsColumnNo[project] - 1] = round(projectsHours[project], 2)
    if projectsHours.has_key(gdocsOtherKey):
        row[otherColumnNo - 1] = round(projectsHours[gdocsOtherKey], 2)

    if commentsColumnNo is not None:
        row[commentsColumnNo - 1] = "\n".join(comments)

    return row


def composeTimeTableForaday(histoyForTheDay, dayStart, workdayLength, truncate):
    """ Iterates through for a single day, and composes a project:hours list from all of that,
    :param histoyForTheDay: a list of the history for a single day
    :param dayStart: start of the day. should have both date and time corresponding to current day start
    :param workdayLength: the work length in hours
    :param truncate: true if the time should be truncated to workday length
    :return: project timetable - a dictionary {projectName:hours}
    """
    history = sorted(histoyForTheDay, key=lambda x: x[3])
    projectsTimetableDict = {}

    currentTime = dayStart
    dayEnd = dayStart + datetime.timedelta(hours = workdayLength)

    # We start at the day start, and move on through the history to each next commit.
    # Then we end by the dayEnd, if truncate is set to True.
    # So we have next periods of commits:
    # (dayStart, commit1), (commit1, commit2), (commit2, ....) [(commitN, dayEnd)]
    # Where the last one only present if truncate == true or commitN time < dayEnd
    lastCommit = None
    for projectName,  rev, changeSet, commitdate, summary in history:
        lastCommit = (projectName,  rev, changeSet, commitdate, summary)
        minutesSpent = commitdate.hour * 60 + commitdate.minute - currentTime.hour * 60 - currentTime.minute
        if projectsTimetableDict.has_key(projectName):
            projectsTimetableDict[projectName] += minutesSpent / 60.0
        else:
            projectsTimetableDict[projectName] = minutesSpent / 60.0
        currentTime = commitdate

    # Check for the last period
    if currentTime < dayEnd and lastCommit is not None: # Assume we were working on the last task the whole day end
        projectName,  rev, changeSet, commitdate, summary = lastCommit
        minutesSpent = dayEnd.hour * 60 + dayEnd.minute - currentTime.hour * 60 - currentTime.minute

        if projectsTimetableDict.has_key(projectName):
            projectsTimetableDict[projectName] += minutesSpent / 60.0
        else:
            projectsTimetableDict[projectName] = minutesSpent / 60.0

    # Truncate the values. We use relatives to truncate that.
    # So if we had 8-hour workday, and we counted 6 hours for project A, and 8 hours for project B,
    # then we assign
    # A = (6 / (6 + 8)) * 8 = 3.42
    # B = (8 / (6 + 8)) * 8 = 4.57
    #                         ____
    #              that gives 8    hours in sum
    if truncate and sum(projectsTimetableDict.values()) > workdayLength:
        originalSum = sum(projectsTimetableDict.values())
        for key in projectsTimetableDict.keys():
            value = projectsTimetableDict[key]
            projectsTimetableDict[key] = (value / originalSum) * workdayLength

    return projectsTimetableDict

availableDatesList = [x for x in worksheet.col_values(dateColumnNo)]
lastAvailableDate = dateutil.parser.parse(availableDatesList[-1])

now = datetime.datetime.now()
today = datetime.datetime(now.year, now.month, now.day)

# Decrease currentNumberOfRows while the indexed row is empty. In this case we find the last filled up row.
currentNumberOfRows = len(availableDatesList)
while availableDatesList[currentNumberOfRows - 1].strip() == "":
    currentNumberOfRows -= 1

# truncate the sheet to the last found date - just to be sure nothing will be left there.
worksheet.resize(rows=currentNumberOfRows)

if today - datetime.timedelta(days=1) > lastAvailableDate:
    date = lastAvailableDate + datetime.timedelta(days=1)

    # Go to the yesterday including.
    while date < today:

        # Compute history for the current date we analyze.
        historyForADate = [(projectName,  rev, changeSet, commitdate, summary)
                           for projectName,  rev, changeSet, commitdate, summary in history
                           if date < commitdate < date + datetime.timedelta(days=1)]
        dayStart = composeADate(date, workdayStart)

        currentDayStart = dayStart
        currentWorkdayLength = workdayLength
        currentTruncate = truncateToWorkday

        projectsHours = composeTimeTableForaday(historyForADate, currentDayStart, currentWorkdayLength,
                                                currentTruncate)

        # This loops infinitely while user do not accept each change. This is done only if silentMode == false
        while True:
            print ""
            print "I choose the following configuration for a day " + str(date) + ". This is " + \
                  WeekNames[date.weekday()]
            print projectsHours

            if silentMode:
                break

            if len(projectsHours) == 0 and date.weekday() < 5:
                print ConsoleColors.Warning + "Warning! Looks like a working day, but no mercurial data was found. " \
                                              "Please check" + ConsoleColors.EndC

            option = raw_input("[A]gree. [E]xit. Change [S]tart of the day. Change workday [L]ength. Change "
                               "[T]runcation to workday rule. [M]anually set a timetable for this day."
                               "Show current [C]onfiguration.").lower()
            # Agree
            if option == 'a':
                break

            # Exit from the application
            if option == 'e':
                exit()

            # Change the start of the day
            if option == 's':
                currentDayStart = parseDayStart(raw_input("Please enter a proposed workday start.").strip())
                currentDayStart = composeADate(date, currentDayStart)

                projectsHours = composeTimeTableForaday(historyForADate, currentDayStart, currentWorkdayLength,
                                                        currentTruncate)

            # Change the work day length
            if option == 'l':
                currentWorkdayLength = int(raw_input("Please enter a proposed workday length.").strip())

                projectsHours = composeTimeTableForaday(historyForADate, currentDayStart, currentWorkdayLength,
                                                        currentTruncate)

            # Change the truncation rule
            if option == 't':
                answer = raw_input("[E]nable or [D]isable the truncation?").strip().lower()
                if answer == 'e':
                    currentTruncate = False
                if answer == 'd':
                    currentTruncate = True

                projectsHours = composeTimeTableForaday(historyForADate, currentDayStart, currentWorkdayLength,
                                                        currentTruncate)

            # Manually set the time table.
            if option == 'm':
                print "Currently supported shortcuts:"
                for prjName, path, usersToFilter, shortcut in projectsData:
                    print "%s as '%s'" % (prjName, shortcut)
                print "%s as '%s'" % (gdocsOtherKey, gdocsOtherShortcut)
                print ""
                print "Example of using shortcuts: 'a3b5'. This will set a project with shortcut a for 3 hours, " \
                      "and a project with shortcut b for 5 hours. Type 'exit' to exit. Type 'back' to go back to " \
                      "selection mode."
                input = raw_input("Please enter your shortcut string").strip()
                if input.lower().strip() == "exit":
                    exit(0)
                if input.lower().strip() == "back":
                    continue

                projectsHours = {}


                # First regex matched each project.
                import re
                for groupMatch in re.match("(\D+\d+)*", input).groups():
                    # This regex splits the matched project into shortcut and the value
                    match = re.match(r"(?P<shortcut>\D+)(?P<value>\d+)", groupMatch)
                    print(match.group("shortcut"))
                    projectsHours[projectShortcutDict[match.group("shortcut")]] = float(match.group("value"))

            # Show current configuraiton
            if option == 'c':
                print "Start day at " + str(currentDayStart)
                print "Day length is " + str(currentWorkdayLength)
                if currentTruncate:
                    print "We truncate to the day length"
                else:
                    print "We do not truncate t the day length"

        comments = ["[%s] %s" % (x[0], x[4]) for x in historyForADate]

        # Finally add the row to the spread sheet
        worksheet.append_row(composeARow(gdocsUserToFill, date, projectsHours, comments))
        currentNumberOfRows += 1
        date += datetime.timedelta(days = 1)

print "Done!"