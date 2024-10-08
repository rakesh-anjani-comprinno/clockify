import dotenv from 'dotenv'
dotenv.config()
import axios from 'axios'
import json2xls from 'json2xls'
import fs from 'fs'
import nodemailer from 'nodemailer'
import moment from 'moment'
import xlsx from 'xlsx'
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const start = "2024-08-01"
const end = "2024-08-31"
const startDateFormat = start + "T00:00:00Z"
const endDateFormat = end + "T23:59:59Z"

const BASE_URL = "https://api.clockify.me/";
const clockifyApiKey = process.env.API_KEY
const headers = {
  'X-Api-Key': clockifyApiKey
};

let params = {
  page_size: 50,
}
const maxRetries = 50;
const fetchDataFromApi = async (url, params, currentRetry = 0) => {
  let results = []
  let page = 1
  while (true) {
    try {
      const response = await axios.get(BASE_URL + url, {
        headers,
        params: { ...params, page: page },
      });
      const responseData = response.data
      results = results.concat(responseData)
      if (responseData.length < params.page_size) {
        break
      }
      page += 1;
    } catch (err) {
      if (err.response.status === 429) {
        const retryAfter = err.response.headers.get('Retry-After') || 1;
        await wait(retryAfter * 1000);
        if (currentRetry < maxRetries) {
          currentRetry++;
          // console.log("currentretry",url,params,"page**",page,currentRetry)
        } else {
          console.log("Max retries reached", page, params, url)
          return err.response.data
        }
      } else {
        console.log("response error data", err.response.data)
        return err.response.data
      }
    }
  }
  return results
};

const wait = (ms) => {
  return new Promise(resolve => setTimeout(resolve, ms));
}

let workspaces = await fetchDataFromApi('api/v1/workspaces', params)

let workspacesIds = workspaces.map((item) => item.id)


let s3DataInJson;
const currentDirectory = __dirname;
const excelFilePath = join(currentDirectory + "/JIRA Tickets.xlsx")
const s3Workbook = xlsx.readFile(excelFilePath)
const s3SheetName = s3Workbook.SheetNames[0]
s3DataInJson = xlsx.utils.sheet_to_json(s3Workbook.Sheets[s3SheetName])
// console.log("sheetData",s3DataInJson)

let usersDataWithWorkspaceId = {};
const usersDataWorkspacePromise = []

const paramsUser = { ...params }
workspacesIds.forEach(workspaceid => {
  usersDataWithWorkspaceId[workspaceid] = []
  usersDataWorkspacePromise.push(fetchDataFromApi(`api/v1/workspaces/${workspaceid}/users`, paramsUser))
});
const allWorkspaceUsers = await Promise.all(usersDataWorkspacePromise)
// console.log(allWorkspaceUsers)

let users = []
let firstUserAdded = false
allWorkspaceUsers.forEach((workspaceUsers, index) => {
  const workspaceId = Object.keys(usersDataWithWorkspaceId)[index];
  usersDataWithWorkspaceId[workspaceId] = workspaceUsers;

  if (firstUserAdded === false) {
    users = workspaceUsers
    firstUserAdded = true
  }

  workspaceUsers.forEach(workspaceUser => {
    const existUser = users.some(user => user.id === workspaceUser.id)
    if (!existUser) {
      users.push(workspaceUser)
    }
  })

});
// console.log("users", users)

// Utility Funtions

const findUserNameById = (userId) => {
  for (const user of users) {
    if (user.id === userId) {
      return user.name
    }
  }
}

const timeInHrsAndMns = (duration) => {
  const hours = Math.floor(duration / (60 * 60 * 1000));
  const minutes = Math.floor((duration % (60 * 60 * 1000)) / (60 * 1000));
  const totalhours = hours + (minutes / 60)
  return totalhours
}

const timeIntervalInMilliseconds = (start, end) => {
  const startTime = new Date(start);
  const endTime = new Date(end);
  return endTime - startTime
}

const convertToMonthFullName = (date) => {
  const [year, monthNumber, day] = date.split("-")
  if (monthNumber >= 1 && monthNumber <= 12) {
    const monthNames = [
      'January', 'February', 'March', 'April',
      'May', 'June', 'July', 'August',
      'September', 'October', 'November', 'December'
    ];
    return monthNames[monthNumber - 1];
  }
}

const formattedDateString = (date) => {
  const [year, month, day] = date.split("-")
  return `${day}-${month}-${year}`
}

const checkTimeEntryInReport = (id) => {
  const timeEntryObj = ticketIdWithEmployeeResults.filter(item => item.TimeEntryId?.includes(id) === true)
  return timeEntryObj
}

const checkTicketIdsInReport = (ticketIds, userName) => {
  const ticketIdObj = ticketIdWithEmployeeResults.filter(item => item.TicketId === ticketIds && item.Name === userName)
  return ticketIdObj
}

const checkAlertStatus = (ticketIdString,summary)=> {
  const isTicketIdMatched = /BSD|LAD|MAD|FAS|NHA|HSD|SAJ|HXA|BASD|NASD|PSAD/.test(ticketIdString)
  const isSummaryMatched = /\[FIRING:1\]/.test(summary)
  const isAlert = isTicketIdMatched + isSummaryMatched
  const alertStatus = isAlert ? 'Alert' : 'Non Alert';
  return alertStatus
}

const checkTicketIdsInTicketIdAndTimeResult = (ticketIds) => {
  const ticketIdObj = ticketIdAndTimeResults.filter(item => item.TicketId === ticketIds)
  return ticketIdObj
}
// Utility Funtion End

// *********Findings the Training Tags********* 
params = {
  ...params,
  start: startDateFormat,
  end: endDateFormat,
}
const tagParams = {
  ...params,
  archived: false
}
const comprinno_workspace_id = workspacesIds[0]
const tagsResult = await fetchDataFromApi(`api/v1/workspaces/${comprinno_workspace_id}/tags`,tagParams)
const trainingObj = tagsResult.filter(e => e.name === "Training")
const trainingId = trainingObj[0].id
// console.log("trainingId",trainingId)


// **********Api call to get all time entries of all user in all workspace and also with training tag ***** *****
const timeEntriesPromises = []
const timeEntriesWithTrainingPromises = []
const traingTagParams = {
  ...params,
  tags:trainingId
}
for (let workspaceId in usersDataWithWorkspaceId) {
  usersDataWithWorkspaceId[workspaceId].forEach((userDoc) => {
    timeEntriesPromises.push(fetchDataFromApi(`api/v1/workspaces/${workspaceId}/user/${userDoc.id}/time-entries`, params))
    timeEntriesWithTrainingPromises.push(fetchDataFromApi(`api/v1/workspaces/${workspaceId}/user/${userDoc.id}/time-entries?`,traingTagParams))
  })
}
const timeEntriesForAllUsers = await Promise.all(timeEntriesPromises)
const timeEntriesWithTrainingTags = await Promise.all(timeEntriesWithTrainingPromises)
// console.log("timeEntriesForAllUser",timeEntriesForAllUsers)
// console.log("timeEntriesWithTrainingTags",timeEntriesWithTrainingTags)

// ********** calculation of two sheets and three worksheet by this big funtion **********
let ticketIdWithEmployeeResults = []
let ticketIdAndTimeResults = []
let ticketIdRegex = /[A-Z]{2,4}-\d{2,4}/g;
let billableHoursInfo = []

for (const timeEntriesForOneUser of timeEntriesForAllUsers) {
  if (timeEntriesForOneUser.length > 0) {
    const userId = timeEntriesForOneUser[0].userId
    const userName = findUserNameById(userId)
    let billableCount = 0
    let billableFlag = true
    let totalBillableHours = 0
    // console.log("userName", userName)
    for (const ticketIdObject of s3DataInJson) {
      const ticketIdToBeFind = ticketIdObject['Issue key']
      const ticketIdToBeFindRegex = new RegExp(`\\b${ticketIdToBeFind}\\b`, 'i')
      let totalTimeInOneTicketIdForOneUser = 0;

      timeEntriesForOneUser.forEach( timeEntry => {
        if(billableFlag){
          if(timeEntry.billable === true){
            const timeForOneTimeEntry = timeIntervalInMilliseconds(timeEntry.timeInterval.start,timeEntry.timeInterval.end)
            const timeInHrsForOneTimeEntry = timeInHrsAndMns(timeForOneTimeEntry)
            totalBillableHours += timeInHrsForOneTimeEntry
          }
        }
        const description = timeEntry.description
        if (ticketIdToBeFindRegex.test(description)) {
          
          const matchedTicketIds = description.match(ticketIdRegex)
          // console.log("matchedTicketIds",matchedTicketIds)
          if(matchedTicketIds?.length > 1) {
            const timeEntryExist = checkTimeEntryInReport(timeEntry.id)
            if(timeEntryExist.length === 0){
              const ticketIdsExist = checkTicketIdsInReport(matchedTicketIds.join(","),userName)
              const timeCombineInMs = timeIntervalInMilliseconds(timeEntry.timeInterval.start,timeEntry.timeInterval.end)
              const timeSpend = timeInHrsAndMns(timeCombineInMs)
              const alertStatus = checkAlertStatus(matchedTicketIds.join(","),ticketIdObject['Summary'])
              if(ticketIdsExist.length === 0){
                ticketIdWithEmployeeResults.push({ Name: userName, TicketId: matchedTicketIds.join(","), Summary: ticketIdObject['Summary'], TimeSpend:timeSpend, "Alert/NonAlert": alertStatus, TimeEntryId:[timeEntry.id]})
                let ticketIdInTicketIdAndTimeResultObj = checkTicketIdsInTicketIdAndTimeResult(matchedTicketIds.join(","))
                if(ticketIdInTicketIdAndTimeResultObj.length === 0){
                  ticketIdAndTimeResults.push({ TicketId: matchedTicketIds.join(","),Summary: ticketIdObject['Summary'],TimeSpend: timeCombineInMs, "Alert/NonAlert": alertStatus })
                }else{
                  ticketIdInTicketIdAndTimeResultObj[0].TimeSpend += timeCombineInMs
                } 
              }else{
                // console.log(userName,ticketIdObject['Issue key'],matchedTicketIds,timeEntry.id,timeSpend,timeEntryExist)
                  for(let item of ticketIdWithEmployeeResults){
                    if(item.TicketId === matchedTicketIds.join(",") && item.Name === userName){
                      item.TimeSpend += timeSpend
                      item.TimeEntryId.push(timeEntry.id)
                    }
                  }
                let ticketIdInTicketIdAndTimeResultObj = checkTicketIdsInTicketIdAndTimeResult(matchedTicketIds.join(","))
                ticketIdInTicketIdAndTimeResultObj[0].TimeSpend += timeCombineInMs
              } 
            }            
          }else {
            const timeDurationStart = timeEntry.timeInterval.start
            const timeDurationEnd = timeEntry.timeInterval.end
            totalTimeInOneTicketIdForOneUser = totalTimeInOneTicketIdForOneUser + timeIntervalInMilliseconds(timeDurationStart, timeDurationEnd)
          }

        }

      })

      billableFlag = false

      if (totalTimeInOneTicketIdForOneUser !== 0) {
        // console.log("issue key",ticketIdObject['Issue key'])
        const alertStatus = checkAlertStatus(ticketIdObject['Issue key'],ticketIdObject['Summary'])
        // console.log("alertstatus",alertStatus)
        // console.log("********************")
        if (ticketIdAndTimeResults.length > 0) {
          const ticketIdAndTimeObj = ticketIdAndTimeResults.filter(item => item.TicketId == ticketIdObject['Issue key'])
          if (ticketIdAndTimeObj.length === 0) {
            const newTicketIdAndTimeObj = { TicketId: ticketIdObject['Issue key'], Summary: ticketIdObject['Summary'], TimeSpend: totalTimeInOneTicketIdForOneUser, "Alert/NonAlert": alertStatus }
            ticketIdAndTimeResults.push(newTicketIdAndTimeObj)
          } else {
            ticketIdAndTimeObj[0].TimeSpend += totalTimeInOneTicketIdForOneUser
          }
        } else {
          const firstTicketIdAndTime = { TicketId: ticketIdObject['Issue key'], Summary: ticketIdObject['Summary'],TimeSpend: totalTimeInOneTicketIdForOneUser,"Alert/NonAlert": alertStatus }
          ticketIdAndTimeResults.push(firstTicketIdAndTime)
        }

        const timeSpend = timeInHrsAndMns(totalTimeInOneTicketIdForOneUser)
        ticketIdWithEmployeeResults.push({ Name: userName, TicketId: ticketIdObject['Issue key'], Summary: ticketIdObject['Summary'], TimeSpend: timeSpend, "Alert/NonAlert": alertStatus })
        // console.log("Results inside TotalTimeInOneTicketIdForOneUser",ticketIdWithEmployeeResults)
      }

    }
    billableHoursInfo.push({userName,totalBillableHours})
  }
}

const trainingAndUsers = []
const totalTrainingTimeCalculate = () => {
  timeEntriesWithTrainingTags.forEach((timeEntriesSet) => {
    if (timeEntriesSet.length > 0) {
      const userId = timeEntriesSet[0].userId
      const userName = findUserNameById(userId)
      let totalTimeInTraining = 0
      timeEntriesSet.forEach(timeEntry => {
        const timeInMs = timeIntervalInMilliseconds(timeEntry.timeInterval.start,timeEntry.timeInterval.end)
        const timeInHrs = timeInHrsAndMns(timeInMs)
        totalTimeInTraining += timeInHrs
      })
      trainingAndUsers.push({userName,totalTimeInTraining})
    }
  })
}
totalTrainingTimeCalculate()

// ********** Converting the time spend in Hours and Minutes for the Sheet1 results [ticketIdAndTimeResults] **********
ticketIdAndTimeResults.forEach((item, index, ticketIdAndTimeResults) => {
  ticketIdAndTimeResults[index].TimeSpend = timeInHrsAndMns(item.TimeSpend)
})
// console.log("ticketIdAndTimeResults",ticketIdAndTimeResults)


// **********Delete the TimeEntry Field in ticketIdWithEmployeeResults**********
for(let item of ticketIdWithEmployeeResults){
  delete item.TimeEntryId
}
// console.log("ticketIdWithEmployeeResults",ticketIdWithEmployeeResults)

// ********* Leave Hour ************
const workingHours = [
  { userName: "Aman Kumar",workday:23},
  { userName: "Aristotle Diogo Fernandes",workday:23},
  { userName: "Mohammed Rizwan",workday:22},
  { userName: "Narahari Mengane",workday:0},
  { userName: "Nikita Dehariya",workday:0},
  { userName: "Parikshit Taksande",workday:23},
  { userName: "Sandeep Malakar",workday:21},
  { userName: "Sandeep Kumar Maurya",workday:22},
  { userName: "Satish Gogiya",workday:21},
  { userName: "Tarun Sharma",workday:0},
  { userName: "Krisnaraj K.C",workday:22},
  { userName: "Pradap V",workday:22},
  { userName: "Atharva Nevase",workday:23},
]

// ********** first worksheet Result in Sheet2 [alert_Non_Alert_result with some extra info] **********
const alertNonAlertResults = []
let currentUser = ticketIdWithEmployeeResults[0].Name
let previousUser = currentUser
let no_of_alert = 0
let no_of_non_alert = 0
let total_time_for_tickets = 0
// let workingHoursofMonths = 21 * 8
for(let i= 0;i<ticketIdWithEmployeeResults.length;i++){
  if(currentUser === ticketIdWithEmployeeResults[i].Name){
      total_time_for_tickets += ticketIdWithEmployeeResults[i].TimeSpend
      if(ticketIdWithEmployeeResults[i]['Alert/NonAlert'] === "Alert"){
          no_of_alert += 1
      }
      if(ticketIdWithEmployeeResults[i]['Alert/NonAlert'] === "Non Alert"){
          no_of_non_alert += 1
      }
      previousUser = currentUser
  }else{
      const billableInfo = billableHoursInfo.filter((info) => info.userName === previousUser )
      const billableHourInMns = billableInfo[0].totalBillableHours 
      
      const trainingInfo = trainingAndUsers.filter((trainingUserDetail) => trainingUserDetail.userName === previousUser)
      let trainingTimeSpend;
      if(trainingInfo.length === 0){
        trainingTimeSpend = 0
      }else{
        trainingTimeSpend = trainingInfo[0].totalTimeInTraining 
      }

      const workInfo = workingHours.filter((info)=> info.userName === previousUser )
      const workHours = workInfo.length ? workInfo[0].workday * 8 : 0

      let memberUtilization = 'N/A'
      if(workHours > 0){
        memberUtilization = ( total_time_for_tickets / (workHours - trainingTimeSpend )) * 100
        memberUtilization =  memberUtilization.toFixed(2) + "%"
      }
      
      alertNonAlertResults.push({"Team member":previousUser,"No.of alert tickets":no_of_alert,"No. of Non alert tickets":no_of_non_alert,"Billable (Hrs)":billableHourInMns,"Total time - Tickets worked on (Hrs)":total_time_for_tickets ,"Training hours":trainingTimeSpend,"Working hours for the month":workHours,"Member Utilization":memberUtilization })
      
      currentUser = ticketIdWithEmployeeResults[i].Name
      total_time_for_tickets = ticketIdWithEmployeeResults[i].TimeSpend
      if(ticketIdWithEmployeeResults[i]['Alert/NonAlert'] === "Alert"){
          no_of_alert = 1
          no_of_non_alert = 0
      }
      if(ticketIdWithEmployeeResults[i]['Alert/NonAlert'] === "Non Alert"){
          no_of_non_alert = 1
          no_of_alert = 0
      }
      previousUser = currentUser
  }
  if(i === ticketIdWithEmployeeResults.length-1 ){
    const billableInfo = billableHoursInfo.filter((info) => info.userName === previousUser )
    const billableHourInMns = billableInfo[0].totalBillableHours 

    const trainingInfo = trainingAndUsers.filter((trainingUserDetail) => trainingUserDetail.userName === previousUser)
    let trainingTimeSpend;
    if(trainingInfo.length === 0){
      trainingTimeSpend = 0
    }else{
      trainingTimeSpend = trainingInfo[0].totalTimeInTraining 
    }

    const workInfo = workingHours.filter((info)=> info.userName === previousUser )
    const workHours = workInfo.length ? workInfo[0].workday * 8 : 0

    let memberUtilization = 'N/A'
    if(workHours > 0){
      memberUtilization = ( total_time_for_tickets / (workHours - trainingTimeSpend )) * 100
      memberUtilization =  memberUtilization.toFixed(2) + "%"
    }

    alertNonAlertResults.push({"Team member":previousUser,"No.of alert tickets":no_of_alert,"No. of Non alert tickets":no_of_non_alert,"Billable (Hrs)":billableHourInMns,"Total time - Tickets worked on (Hrs)":total_time_for_tickets,"Training hours":trainingTimeSpend,"Working hours for the month":workHours,"Member Utilization":memberUtilization })
  }
}
// console.log(alertNonAlertResults)


// ********** Generating First Sheet 1 **********
const ticketIdAndTimeWorkbook = xlsx.utils.book_new();
const ticketIdAndTimeResultsWorksheet = xlsx.utils.json_to_sheet(ticketIdAndTimeResults);
xlsx.utils.book_append_sheet(ticketIdAndTimeWorkbook, ticketIdAndTimeResultsWorksheet,'Sheet1');
const ticketIdAndTimeXls = xlsx.write(ticketIdAndTimeWorkbook, {bookType:'xlsx', type: 'buffer'})
const ticketIdAndTimeFileName = `Sheet1_${formattedDateString(start)} to ${formattedDateString(end)}.xlsx`
// xlsx.writeFile(ticketIdAndTimeWorkbook, ticketIdAndTimeFileName );

// ********** Generating Second Sheet 2 **********
const ticketIdWithEmployeeWorkbook = xlsx.utils.book_new();
const ticketIdWithEmployeeResultsWorksheet = xlsx.utils.json_to_sheet(ticketIdWithEmployeeResults);
const alertNonAlertResultsWorksheet = xlsx.utils.json_to_sheet(alertNonAlertResults);
xlsx.utils.book_append_sheet(ticketIdWithEmployeeWorkbook, alertNonAlertResultsWorksheet, 'Summary');
xlsx.utils.book_append_sheet(ticketIdWithEmployeeWorkbook, ticketIdWithEmployeeResultsWorksheet, 'Sheet2');
const ticketIdWithEmployeeXls = xlsx.write(ticketIdWithEmployeeWorkbook, { bookType: 'xlsx', type: 'buffer' });
const ticketIdWithEmployeeFileName = `Sheet2_${formattedDateString(start)} to ${formattedDateString(end)}.xlsx`
// xlsx.writeFile(ticketIdWithEmployeeWorkbook, ticketIdWithEmployeeFileName)


// ********** Sending Mail  **********
const sendEmail = async (emails) => {
  const transport = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.SMTP_USERNAME,
      pass: process.env.SMTP_PASSWORD,
    }
  })

  const mailOptions = {
    from: process.env.SMTP_USERNAME,
    to: emails,
    subject: `Managed Services Team Timesheet Reports: ${convertToMonthFullName(start)} ${start.split("-")[0]}`,
    text: `
    Please find attached the timesheet reports for our Managed Services team.
    • The Sheet 1 provides a comprehensive monthly overview, detailing the total time spent on each ticket.
    • The Sheet 2 offers weekly updates on team members' time entries, including ticket numbers and hours worked
    `,
    attachments: [
      {
        filename: ticketIdAndTimeFileName,
        content: ticketIdAndTimeXls,
        encoding: 'base64',
      },
      {
        filename: ticketIdWithEmployeeFileName,
        content: ticketIdWithEmployeeXls,
        encoding: 'base64',
      },
    ]
  }
  await transport.sendMail(mailOptions)
  console.log('mail send.....')
}
const emails = ['ankith.s@comprinno.net']
// const emails = ['rkanjani14@gmail.com']
// const emails = ['ankith.s@comprinno.net','coe@comprinno.net']

sendEmail(emails)
