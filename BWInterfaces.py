import sqlite3
import xlsxwriter
from collections import namedtuple

__author__ = 'Jon Waterhouse'
class BWSchedules():
    def calculate_interfaces(self,path,db):
        sqlite_db = path + db
        conn = sqlite3.connect(sqlite_db)
        c = conn.cursor()
        c.executescript("""
            DROP VIEW IF EXISTS SchRefBW;

            CREATE VIEW SchRefBW AS
                   SELECT DISTINCT s.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM schedule AS s
                                   INNER JOIN sch_all AS c
                                           ON s.schedule = c.schedule
                             WHERE ( c.line LIKE "%GDW%"
                                              OR
                                              c.line LIKE "% BW %"
                                                     OR
                                                     c.line LIKE "%P2W%"
                                                            OR
                                                        c.line LIKE "%P4W%" )
                             ORDER BY s.schedule;

            DROP VIEW IF EXISTS [SchRefBW-IF];

            CREATE VIEW [SchRefBW-IF] AS
                   SELECT DISTINCT s.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM SchRefBW AS s
                                   INNER JOIN sch_lines AS l
                                           ON s.schedule = l.schedule
                                   INNER JOIN jobs AS j
                                           ON l.job = j.job
                             WHERE ( ( s.Platform LIKE "SP_W-023%"
                                                      AND
                                                  j.platform NOT LIKE "SP_W-023%" )
                                       OR
                                       ( s.Platform NOT LIKE "SP_W-023%"
                                                          AND
                                                      j.platform LIKE "SP_W-023%" )
                                       OR
                                       ( s.Platform NOT LIKE "SP_W-023%"
                                                          AND
                                                  j.platform NOT LIKE "SP_W-023%" )  )
                             GROUP BY s.schedule,
                                      j.job;

            DROP VIEW IF EXISTS SchBWOpensFile;

            CREATE VIEW SchBWOpensFile AS
                   SELECT DISTINCT o.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM sch_opens AS o
                                   INNER JOIN schedule AS s
                                           ON o.schedule = s.schedule
                             WHERE o.schedule LIKE "GDW%"
                             GROUP BY o.schedule;

            DROP VIEW IF EXISTS SchBWLinked;

            CREATE VIEW SchBWLinked AS
                   SELECT DISTINCT l.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM sch_links AS l
                                   INNER JOIN schedule AS s
                                           ON l.schedule = s.schedule
                             WHERE l.schedule LIKE "%GDW%"
                                              AND
                                              l.precedes NOT LIKE "%GDW%"
                                                             OR
                                                             l.schedule NOT LIKE "%GDW%"
                                                                            AND
                                                                            l.precedes LIKE "%GDW%"
                             GROUP BY l.schedule;

            DROP VIEW IF EXISTS SchBWNeedsNonBWRsrc;

            CREATE VIEW SchBWNeedsNonBWRsrc AS
                   SELECT DISTINCT n.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM sch_needs AS N
                                   INNER JOIN schedule AS s
                                           ON n.schedule = s.schedule
                             WHERE n.schedule LIKE "%GDW%"
                                              AND
                                              n.needs NOT LIKE "%GDW%"
                                                          OR
                                                          n.schedule NOT LIKE "%GDW%"
                                                                         AND
                                                                         n.needs LIKE "%GDW%"
                             GROUP BY n.schedule;

            DROP VIEW IF EXISTS SchToFromBWExtract;

            CREATE VIEW SchToFromBWExtract AS
                   SELECT DISTINCT l.schedule,
                                   s.name,
                                   s.platform,
                                   s.[ACTION]
                              FROM schedule AS s
                                   INNER JOIN sch_all AS l
                                           ON l.schedule = s.schedule
                             WHERE ( l.schedule LIKE "%GDW%"
                                                  AND
                                              l.line LIKE "%Extract%" )
                                   OR
                                   ( l.schedule NOT LIKE "%GDW%"
                                                      AND
                                                  l.line LIKE "%GDW%" )
                             GROUP BY l.schedule;

            DROP VIEW IF EXISTS BWIFSchedules;

            CREATE VIEW BWIFSchedules AS
                   SELECT *
                     FROM SchBWLinked
                   UNION
                   SELECT *
                     FROM SchBWNeedsNonBWRsrc
                   UNION
                   SELECT *
                     FROM SchBWOpensFile
                   UNION
                   SELECT *
                     FROM [SchRefBW-IF]
                   UNION
                   SELECT *
                     FROM SchToFromBWExtract;

            DROP VIEW IF EXISTS BWIFSchedsWithRuntimes;

            CREATE VIEW BWIFSchedsWithRuntimes AS
                   SELECT B.schedule, B.name, b.platform, b.action, f.freq
                     FROM BWIFSchedules AS B
                          INNER JOIN SCH_FREQ AS F
                                  ON B.schedule = F.schedule
                    ORDER BY B.schedule;
        """)
        conn.commit()

    def output(self, path, db, output_loc):
        """
        use xlsxwriter module to write database view to excel.
        """
        #Find calendars which we are not interested in
        cal_file_name = 'PROD-CALENDARS-10122015.txt'
        interesting_days = ['11/13/2015', '11/14/2015','11/15/2015','11/16/2015','11/17/2015','11/18/2015']
        cals_to_exclude = self.find_calendars_with_dates(path,cal_file_name,interesting_days)
        workbook = xlsxwriter.Workbook(path + output_loc)
        worksheet = workbook.add_worksheet()
        #Set up display format
        worksheet.set_zoom(77)
        worksheet.set_column(0,0,20)
        worksheet.set_column(1,1,60)
        worksheet.set_column(2,2,15)
        worksheet.set_column(3,3,20)
        worksheet.set_column(4,4,40)
        worksheet.set_column(5,5,15)
        worksheet.set_column(6,6,60)
        worksheet.set_column(7,7,60)
        #Column headings
        format = workbook.add_format({'bold':True})
        worksheet.write(0,0,'SCHEDULE',format)
        worksheet.write(0,1,'NAME',format)
        worksheet.write(0,2,'PLATFORM',format)
        worksheet.write(0,3,'ACTION',format)
        worksheet.write(0,4,'FREQ',format)
        worksheet.write(0,5,'POSS RELVT CAL', format)
        worksheet.write(0,6,'PRECEDES',format)
        worksheet.write(0,7,'FOLLOWS',format)
        worksheet.freeze_panes(1,0)
        #Read the database to get the relevant data
        sqlite_db = path + db
        conn = sqlite3.connect(sqlite_db)
        c = conn.cursor()
        record = namedtuple('record', 'SCHEDULE NAME PLATFORM ACTION FREQ NONREL')
        rows = []
        # Output of next statement is schedule, name, platform, action, freq
        for row in c.execute('SELECT * FROM [BWIFSchedsWithRuntimes] ORDER BY SCHEDULE'):
            if row[4].startswith('RUNCYCLE CALENDAR'):
                cal = row[4][row[4].rfind(' ')+1:len(row[4])].strip('\n')
                if cal in cals_to_exclude: relevant = 'N'
                else: relevant = 'Y'
            else: relevant = 'Y'
            if row[4].startswith('REQUEST'): relevant = 'N'
            rows.append(record(row[0], row[1],row[2],row[3],row[4],relevant))
        #Read database to get preceding and following schedules and write lines to excel
        i,j = 1, 0
        for record in rows:
            precedes, follows = [],[]
            for prec in c.execute("SELECT PRECEDES FROM SCH_LINKS WHERE SCHEDULE = ? ORDER BY PRECEDES", (record.SCHEDULE,)):
                precedes.append(prec[0])
            for foll in c.execute("SELECT SCHEDULE FROM SCH_LINKS WHERE PRECEDES = ? ORDER BY SCHEDULE",(record.SCHEDULE,)):
                follows.append(foll[0])
            for el in record:
                worksheet.write(i,j,el)
                j += 1
            worksheet.write(i,j,', '.join(precedes)) #All preceding to fit in one cell
            worksheet.write(i,j+1,', '.join(follows)) #All following to fit in one cell
            i += 1 # Increment row index
            j = 0 # Reset col index
        #Now we know the last row number apply an autofilter
        worksheet.autofilter(0,0,i-1,6)

    def find_calendars_with_dates(self,path, cal_file_name, interesting_days):
        """
        Read calendar, identify dates they contain and find calendars not containing specific dates
        :param path: Path to calendar file
        :param cal_file_name: calendar file name
        :param interesting_days: those days we want to see if a calendar contains. A list of dates in US format
        :return: list of calendars not containing the interesting dates
        """
        cal_file = open(path + cal_file_name, 'r')
        calendar = {}
        cal_dates = []
        for line in cal_file:
            if line.startswith('$CALENDAR'): continue
            if line.startswith('  "'): continue
            if not line.startswith(' ') and not line.startswith('\n'):
                calendar_name = line[0:line.find(' ')]
                cal_dates = []
            if line.startswith(' ') and line != '':
                cal_dates.append(line.rstrip('\n').strip())
            if line == '\n': # End calendar
                calendar[calendar_name] = cal_dates
                continue
        s = set()
        for k in calendar.keys():
            #calendar[k] is a list l=each elemsnt contains 'mm/dd/yy mm/dd/yy ...'
            calendar[k] = ' '.join(calendar[k]).split(' ')
            interesting = False
            for d in interesting_days:
                if d in calendar[k]: interesting = True
            if not interesting: s.add(k) # remove duplicates
        for el in s: print (el,calendar[el])
        return list(s)


if __name__ == '__main__':
    path = 'c:\\Users\\u104675\\Jon_Waterhouse_Docs\\OneDrive - Eastman Koda~1\\PythonProjects\\Maestro\\'
    db = 'schedule.db'
    s = BWSchedules()
    s.calculate_interfaces(path,db)
    output_xlsx = 'BW Interface Schedules.xlsx'
    s.output(path, db, output_xlsx)

    #s.find_calendars_with_dates(path, cal_file_name, interesting_days)