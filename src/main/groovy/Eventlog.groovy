/**
 * Created by Michael Rinus (michael.rinus@holisticsystems.de ) on 14.12.2016.
 *
 *
 * This script is an adapted version of an example i found relating Scriptom querying the Windows Eventlog.
 *
 * The initial idea was "How can i approximately find out from when to when i worked on a certain Windows-PC
 * at specific days" - to back up my time estimation for customers.
 *
 * It's actually a bit random, at least on the PC where i use it, but fine enough to get a good estimation.
 *
 * As a result a timetable is writen for each day, including start and end time plus calculating the brutto and netto
 * workhours using a simple rule for pause times.
 *
 * Feel free to do with it what you want!
 *
 */
import org.codehaus.groovy.scriptom.ActiveXObject
import org.codehaus.groovy.scriptom.Scriptom

import java.text.SimpleDateFormat

//System.setProperty("org.codehaus.groovy.scriptom.debug", "true")

final String localTimeZone = "GMT+01:00"

Scriptom.inApartment
        {
            def strComputer = "."

            def objWMIService = new ActiveXObject(
                    "winmgmts:{impersonationLevel=impersonate,(Security)}!\\\\$strComputer\\root\\cimv2"
            )

            //Filtering for a specific Start of TimeWritten as the Eventlog gets quite large
            def filter = [year: '2016', month: '12']

            //Get Events from EventLog
            def colLogEvents = objWMIService.invokeMethod("ExecQuery", "Select * from Win32_NTLogEvent Where Logfile = 'System' and TimeWritten > '${filter.year}${filter.month}01000000.000000-000'")

            //Converts the String written into the EventLog to a usable type considering possible TimeZone issues which i faced.
            def timeWrittenToDate = { String timeWritten ->
                // Example Time from EventLog: "20161101000000.000000-000"
                // Time seems to be UTC in a String - so keep in mind transform it to your local TZ

                def example = "20161025170848.308659-000"

                final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss")
                sdf.setTimeZone(TimeZone.getTimeZone("GMT"))
                def exampleDate = sdf.parse(example)

                assert exampleDate instanceof Date && exampleDate == Date.parse('dd.MM.yyyy HH:mm:ss', '25.10.2016 19:08:48')

                sdf.parse(timeWritten)
            }

            //Collect the timeWritten in a list
            def timeList = colLogEvents.collect { timeWrittenToDate(it.TimeWritten) }

            /*
               I left this in so you can see how to use the other properties without searching the web

               colLogEvents.each{
               //Collect the timeWritten in a list
               timeList << timeWrittenToDate(it.TimeWritten)
               def entry = [timeWrittenToDate(it.TimeWritten), it.Category, it.ComputerName , it.EventCode, it.Message, it.RecordNumber, it.SourceName, it.Type, it.User ]
               println entry
            } */

            //Sort out dupes
            timeList = timeList.sort().unique()

            //A Map where the Date as "dd.MM.yyyy" is the key
            LinkedHashMap timeMap = [:]

            Calendar localDate = new GregorianCalendar(TimeZone.getTimeZone(localTimeZone))

            //To build the key used in the timeMap
            def justTheDate = { Calendar someDate ->
                someDate?.format('dd.MM.yyyy')
            }

            timeList.each { date ->
                localDate.setTimeInMillis(date.time)
                //If there is anything stored in timeMap[(justTheDate(localDate))] like timeMap['01.01.2016] then add the value [localDate.time] to the List
                //Otherwise initialize the List with [localDate.time]
                timeMap[(justTheDate(localDate))] = timeMap[(justTheDate(localDate))] ? timeMap[(justTheDate(localDate))] + [localDate.time] : [localDate.time]
            }

            //Calculate min and max time for each Map-Ley and replace its value
            timeMap.each { entry ->
                entry.with { value = [min: value.min(), max: value.max()] }
            }

            //Header
            println(['Day'.center(9), 'Start'.padLeft(6), 'End'.center(5), 'Pause', 'brt.Worktime', 'net.Worktime'].join(' '))

            timeMap.each { entry ->
                def millisToHours = 3600000.0
                def day = entry.key

                def start = entry.value.min
                def end = entry.value.max

                //Decimal hours for brutto work time
                def brtWorkTime = { (((end.time - start.time) / millisToHours) as Float).round(2) }

                //Pause calculation as 30 Minutes for worktime less than 9 hours and 45 Minutes for more
                def pause = { time, threshold = 9, smallPause = 0.50, bigPause = 0.75 ->
                    (time >= threshold) ? bigPause : smallPause
                }

                //A bit spice
                pause = pause.curry(brtWorkTime())

                def brt2Net = { time -> (time - pause(time)).round(2) }

                //Chain the Closures to make the last statement more readable - or more confusing...
                def netWorkTime = brt2Net << brtWorkTime

                def line = [day, start.format('HH:mm'), end.format('HH:mm'), pause().toString().padLeft(5), brtWorkTime().toString().padLeft(12), netWorkTime().toString().padLeft(12)]

                println line.join(' ')
            }
            println 'Script is done.'
        }
