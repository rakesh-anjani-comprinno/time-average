import axios from 'axios'
import dotenv from 'dotenv'
dotenv.config()
import xlsx from 'xlsx'
import Excel from 'exceljs'

// Some Configuration for Fetching Api
const BASE_URL = 'https://comprinno-tech.atlassian.net'
const username = process.env.JIRA_USERNAME
const password = process.env.JIRA_PASSWORD
const headers = {
    'Authorization': 'Basic ' + btoa(username + ':' + password)
};

// Api Call made
const fetchDataFromApi = async (url, params) => {
    try {
        const response = await axios.get(BASE_URL + url, {
            headers,
            params: { ...params },
        });
        return response.data
    } catch (err) {
        console.log("response error data", err)
        return err
    }
}

// Time Duration
const startDate = '2024-07-01'
const endDate = '2024-08-01'


// Clients
const clients = ['Boat','HivePro','Comprinno','Fiery','Geotree','Lightmetrics','Mantle Labs','NeuralHive','Tevico','Sirovate','Billeasy','Propertyshare','Neuralbits']

// Alert - Non Alert Funtion
const checkAlertStatus = (key,summary)=> {
    const isTicketIdMatched = /BSD|LAD|MAD|FAS|NHA|HSD|HXA|SAJ|BASD|NASD|PSAD/.test(key)
    const isSummaryMatched = /\[FIRING:1\]/.test(summary)
    const isAlert = isTicketIdMatched + isSummaryMatched
    // const alertStatus = isAlert ? 'Alert' : 'Non Alert';
    const alertStatus = isAlert ? true : false;

    return alertStatus
  }

const timeInHrsAndMins = (time) => {
    const timeRegex = /(?:(\d+)h\s*)?(?:(\d+)m)?/;
    const regexMatch = time.match(timeRegex);
    const hours = regexMatch[1] ? parseInt(regexMatch[1]) : 0;
    const minutes = regexMatch[2] ? parseInt(regexMatch[2]) : 0;
    return { hours, minutes }
}

const sumOfTime = ( firstTimeString, secondTimeString) => {
    const { hours:firstTimeHours, minutes:firstTimeMinutes } = timeInHrsAndMins(firstTimeString)
    const { hours:secondTimeHours , minutes:secondTimeMinutes } = timeInHrsAndMins(secondTimeString)

    let totalHours = firstTimeHours + secondTimeHours
    let totalMinutes = firstTimeMinutes + secondTimeMinutes

    if(totalMinutes > 60){
        totalHours +=  Math.floor(totalMinutes / 60)
        totalMinutes = totalMinutes % 60
    }
    return { hours: totalHours, minutes:totalMinutes }
}

const avgOfTime = (hr,min,count) => {
    const totalMinutes = hr * 60 + min 
    const avgMinutes = Math.floor(totalMinutes / count)
    const hours = Math.floor(avgMinutes / 60)
    const minutes = avgMinutes % 60
    // console.log("avgofTime :hours",hours," minutes: ",minutes)
    const avgTimeColonString = `${hours}:${minutes}`
    const avtTimeObjForm = {hours,minutes}
    return { avgTimeColonString, avtTimeObjForm }
}

function subtractTime(time1, time2) {
    const {hours: hours1, minutes: minutes1} = timeInHrsAndMins(time1);
    const {hours: hours2, minutes: minutes2} = timeInHrsAndMins(time2);

    // Convert both times to minutes
    const totalMinutes1 = hours1 * 60 + minutes1;
    const totalMinutes2 = hours2 * 60 + minutes2;

    // Perform subtraction
    const differenceMinutes = totalMinutes1 - totalMinutes2;

    // Convert the difference back to hours and minutes
    const hoursDifference = Math.floor(differenceMinutes / 60);
    const minutesDifference = differenceMinutes % 60;

    return `${hoursDifference}h ${minutesDifference}m`
}

const additionTime = (time1,time2) => {
    const {hours: hours1, minutes: minutes1} = timeInHrsAndMins(time1);
    const {hours: hours2, minutes: minutes2} = timeInHrsAndMins(time2);

    // Convert both times to minutes
    const totalMinutes1 = hours1 * 60 + minutes1;
    const totalMinutes2 = hours2 * 60 + minutes2;

    // Perform subtraction
    const differenceMinutes = totalMinutes1 + totalMinutes2;

    // Convert the difference back to hours and minutes
    const hoursDifference = Math.floor(differenceMinutes / 60);
    const minutesDifference = differenceMinutes % 60;

    return `${hoursDifference}h ${minutesDifference}m`
}

const isNegativeTime = (timeString) => {
    return timeString.startsWith("-");
}

const formattedTime = (ms) => {
    const totalSeconds = Math.floor(ms/1000)
    const totalMinutes = Math.floor( totalSeconds / 60 )
    const seconds = totalSeconds % 60
    const totalHours = Math.floor(totalMinutes / 60)
    const minutes = totalMinutes % 60
    return `${totalHours}:${minutes}:${seconds}`
}

const timeSpendInMilliSeconds = (timeSpend) => {
    const [ hr, min, sec ] = timeSpend.split(":")
    const totalMilliseconds = ( hr * 3600 + min * 60 + sec) * 1000
    return totalMilliseconds
}

const jiraIssuesSearch = async (startDate,endDate,dateType) => {
    let data;
    let params = {
        jql: `${dateType} >= '${startDate}' AND ${dateType} <= '${endDate}' ORDER BY ${dateType} DESC`,
        startAt:0,
        maxResults:100,
        fields:'summary,project,priority,customfield_10094,customfield_10092'
    }
    let results = [];
    while(true){
       const data = await fetchDataFromApi('/rest/api/3/search', params)
       results = [...results,...data.issues]
    //    console.log("startAt",data.startAt)
       console.log("data",data)
       if(data.issues.length < data.maxResults){
        break
       }
       params.startAt = data.startAt + data.maxResults
    }
    // console.log("results",results)
    console.log("break*********************break")
    return results 
}

// Initializing all clients
const clientsAllTimes = clients.map(client => {

    return {
        client,
        avgResponseTime: { 
            alerts: { critical: { issueCount: 0, spendTime:"0m" }, high: { issueCount: 0, spendTime:"0m" }, medium: { issueCount: 0, spendTime:"0m" }, low: { issueCount: 0, spendTime:"0m" } },
            nonAlerts: { critical:{ issueCount: 0, spendTime:"0m" }, high:{ issueCount: 0, spendTime:"0m" }, medium:{ issueCount: 0, spendTime:"0m" }, low:{ issueCount: 0, spendTime:"0m" } }
        },
        avgResolutionTime: {
            alerts: { critical: { issueCount: 0, spendTime:"0m" }, high: { issueCount: 0, spendTime:"0m" }, medium: { issueCount: 0, spendTime:"0m" }, low: { issueCount: 0, spendTime:"0m" } },
            nonAlerts: { critical:{ issueCount: 0, spendTime:"0m" }, high:{ issueCount: 0, spendTime:"0m" }, medium:{ issueCount: 0, spendTime:"0m" }, low:{ issueCount: 0, spendTime:"0m" } }
        }
    }
})

// calculating total Time spend by each client on each seperation
const calculateClientsAllTimesResponse = async() => {
    const jiraResponseIssues = await jiraIssuesSearch(startDate,endDate,'created')
    jiraResponseIssues.forEach((issue) => {
    
        const key = issue.key
        const summary = issue.fields.summary
        const isAlert = checkAlertStatus(key,summary)
        const projectName = issue.fields.project.name
        // console.log("Response key: ",issue)
        const mappedClient = clients.find((client) => {
            if(projectName.toLowerCase().includes(client.toLowerCase())){
                return client
            }
            if(projectName==='Mantle lab Alerts Service desk ' && client === 'Mantle Labs'){
                return client
            }
            if(projectName==='Neural Hive Alert Jira ' && client === 'NeuralHive'){
                return client
            }
        })
        // console.log("mappedClient",mappedClient,projectName)
        const avgTimeClientObj = clientsAllTimes.filter((clientDetail)=> clientDetail.client === mappedClient)
        // console.log("issue priority",issue)
        const priority = issue.fields.priority.name
    
        // Finding response time
        const responseTime = issue.fields?.customfield_10094
        if(responseTime?.completedCycles?.length){
            const goalDuration = responseTime.completedCycles[0].goalDuration.friendly
            const remainingTime = responseTime.completedCycles[0].remainingTime.friendly
            const isNegativeTimePresent = isNegativeTime(remainingTime)
            let timespend;
            if(isNegativeTimePresent){
                const remainingTimePositive = remainingTime.replace("-","");
                timespend = additionTime(goalDuration,remainingTimePositive)
            }else{
                timespend = subtractTime(goalDuration,remainingTime)
            }

            // console.log("responseTIme,",timespend)
            // console.log("responseTIme,",avgTimeClientObj)
            // console.log("Key",key,"isAlert",isAlert,"priority",priority,"projectname",projectName,"mappedClient",mappedClient,"responseTime",timespend)
            // if( (key.includes("LAD") || key.includes("LM")) && priority === 'High'){
            //     console.log("LAD LM issues",key,JSON.stringify(avgTimeClientObj,null,2))
            // }
            if(avgTimeClientObj.length){
                if(isAlert){
    
                    if(priority.toLowerCase() === 'highest'){
                        const spendTime = avgTimeClientObj[0].avgResponseTime.alerts.critical.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.alerts.critical.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.alerts.critical.issueCount += 1
    
                    }
                    else if(priority.toLowerCase() === 'high'){
                        
                        const spendTime = avgTimeClientObj[0].avgResponseTime.alerts.high.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.alerts.high.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.alerts.high.issueCount += 1

                    }
                    else if(priority.toLowerCase() === 'medium'){
                        
                        const spendTime = avgTimeClientObj[0].avgResponseTime.alerts.medium.spendTime
                        // console.log("Medium spendTime, timespend",spendTime,timespend)
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.alerts.medium.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.alerts.medium.issueCount += 1
                        

                    }
                    else if(priority.toLowerCase() === 'low'){
                        // console.log("key: ",key,"priority: ",priority,"isAlert: ",isAlert,"response: ",timespend)

                        const spendTime = avgTimeClientObj[0].avgResponseTime.alerts.low.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.alerts.low.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.alerts.low.issueCount += 1
                
                        // console.log(avgTimeClientObj[0])
                    }
    
                }else{
                    if(priority.toLowerCase() === 'highest'){
                        const spendTime = avgTimeClientObj[0].avgResponseTime.nonAlerts.critical.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.critical.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.critical.issueCount += 1

                    }
                    else if(priority.toLowerCase() === 'high'){

                        const spendTime = avgTimeClientObj[0].avgResponseTime.nonAlerts.high.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.high.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.high.issueCount += 1

                    }
                    else if(priority.toLowerCase() === 'medium'){
                        // console.log("priority: ",priority,"timeSpend: ",timespend,"isAlert: ",isAlert)

                        const spendTime = avgTimeClientObj[0].avgResponseTime.nonAlerts.medium.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.medium.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.medium.issueCount += 1

                        // console.log(JSON.stringify(avgTimeClientObj[0]))
                    }
                    else if(priority.toLowerCase() === 'low'){

                        const spendTime = avgTimeClientObj[0].avgResponseTime.nonAlerts.low.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.low.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResponseTime.nonAlerts.low.issueCount += 1

                    }
                }
            }
        }
    })

}
await calculateClientsAllTimesResponse()

// calculating total Time spend by each client on each Resolution
const calculateClientsAllTimesResolution = async() => {
    const jiraResoutionIssues = await jiraIssuesSearch(startDate,endDate,"updated")
    jiraResoutionIssues.forEach((issue) => {
    
        const key = issue.key
        const summary = issue.fields.summary
        const isAlert = checkAlertStatus(key,summary)
        const projectName = issue.fields.project.name
        // console.log("Resoution key: ",key)
        const mappedClient = clients.find((client) => {
            if(projectName.toLowerCase().includes(client.toLowerCase())){
                return client
            }
            if(projectName==='Mantle lab Alerts Service desk ' && client === 'Mantle Labs'){
                return client
            }
            if(projectName==='Neural Hive Alert Jira ' && client === 'NeuralHive'){
                return client
            }
        })
        // console.log("mappedClient",mappedClient,projectName)
        const avgTimeClientObj = clientsAllTimes.filter((clientDetail)=> clientDetail.client === mappedClient)
        // console.log("issue priority",issue)
        const priority = issue.fields.priority.name
    
        // Findings resolution time
        const resolutionTime = issue.fields?.customfield_10092
        if(resolutionTime?.completedCycles?.length){
            // const timespend = resolutionTime.completedCycles[0].elapsedTime.friendly

            const goalDuration = resolutionTime.completedCycles[0].goalDuration.friendly
            const remainingTime = resolutionTime.completedCycles[0].remainingTime.friendly
            const isNegativeTimePresent = isNegativeTime(remainingTime)
            let timespend;
            if(isNegativeTimePresent){
                const remainingTimePositive = remainingTime.replace("-","");
                timespend = additionTime(goalDuration,remainingTimePositive)
                // if(mappedClient === "Boat" && priority === 'High' && isAlert=== false){
                //     console.log("boat nonalert high timepsend",timespend)
                //     console.log("boat nonalert high timepsend someother details",goalDuration,remainingTime,remainingTimePositive,timespend)
                // }
            }else{
                timespend = subtractTime(goalDuration,remainingTime)
            }
            
            // console.log("Key",key,"isAlert",isAlert,"priority",priority,"projectname",projectName,"mappedClient",mappedClient,"resolutionTime",timespend)
            if(avgTimeClientObj.length){
                if(isAlert){
                    if(priority.toLowerCase() === 'highest'){

                        const spendTime = avgTimeClientObj[0].avgResolutionTime.alerts.critical.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.alerts.critical.spendTime = hours + "h " + minutes + "m"
                        avgTimeClientObj[0].avgResolutionTime.alerts.critical.issueCount += 1
                    }
                    else if(priority.toLowerCase() === 'high'){

                        const spendTime = avgTimeClientObj[0].avgResolutionTime.alerts.high.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.alerts.high.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.alerts.high.issueCount += 1
                    }
                    else if(priority.toLowerCase() === 'medium'){

                        const spendTime =avgTimeClientObj[0].avgResolutionTime.alerts.medium.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.alerts.medium.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.alerts.medium.issueCount += 1
                    }
                    else if(priority.toLowerCase() === 'low'){

                        // console.log("key: ",key,"priority: ",priority,"isAlert: ",isAlert,"resolution: ",timespend)
                        const spendTime = avgTimeClientObj[0].avgResolutionTime.alerts.low.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.alerts.low.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.alerts.low.issueCount += 1

                        // console.log(avgTimeClientObj[0])
                    }
    
                }else{
                    if(priority.toLowerCase() === 'highest'){
                        const spendTime = avgTimeClientObj[0].avgResolutionTime.nonAlerts.critical.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.critical.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.critical.issueCount += 1
                    }
                    else if(priority.toLowerCase() === 'high'){
                        
                        const spendTime = avgTimeClientObj[0].avgResolutionTime.nonAlerts.high.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.high.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.high.issueCount += 1
                        // if(mappedClient==="Boat"){
                        //     console.log("avgTImelcientObj of Boat",JSON.stringify(avgTimeClientObj))
                        // }
                    }
                    else if(priority.toLowerCase() === 'medium'){
                        const spendTime = avgTimeClientObj[0].avgResolutionTime.nonAlerts.medium.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.medium.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.medium.issueCount += 1
                    }
                    else if(priority.toLowerCase() === 'low'){
                        const spendTime = avgTimeClientObj[0].avgResolutionTime.nonAlerts.low.spendTime
                        const {hours , minutes} = sumOfTime(spendTime,timespend)
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.low.spendTime = hours + "h " + minutes + "m"
                        
                        avgTimeClientObj[0].avgResolutionTime.nonAlerts.low.issueCount += 1

                        // if(mappedClient==="HivePro"){
                        //     console.log()
                        //     console.log("avgTImelcientObj of HivePro",JSON.stringify(avgTimeClientObj))
                        // }
                    }
                }
            }
        }
    })

}
await calculateClientsAllTimesResolution()

// Adding Summary Obj
const addSummaryObj = () => {
    clientsAllTimes.push({
        client:"Summary",
        avgResponseTime: { 
            alerts: { critical: { issueCount: 0, spendTime:"0m" }, high: { issueCount: 0, spendTime:"0m" }, medium: { issueCount: 0, spendTime:"0m" }, low: { issueCount: 0, spendTime:"0m" } },
            nonAlerts: { critical:{ issueCount: 0, spendTime:"0m" }, high:{ issueCount: 0, spendTime:"0m" }, medium:{ issueCount: 0, spendTime:"0m" }, low:{ issueCount: 0, spendTime:"0m" } }
        },
        avgResolutionTime: {
            alerts: { critical: { issueCount: 0, spendTime:"0m" }, high: { issueCount: 0, spendTime:"0m" }, medium: { issueCount: 0, spendTime:"0m" }, low: { issueCount: 0, spendTime:"0m" } },
            nonAlerts: { critical:{ issueCount: 0, spendTime:"0m" }, high:{ issueCount: 0, spendTime:"0m" }, medium:{ issueCount: 0, spendTime:"0m" }, low:{ issueCount: 0, spendTime:"0m" } }
        }
    })
}
addSummaryObj()

const excelSheetInitialization = () => {

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    worksheet.mergeCells('A1:A3');
    worksheet.getCell('A1').value = 'Customer';

    // Merge cells for "Response Time"
    worksheet.mergeCells('B1:K1');
    worksheet.getCell('E1').value = 'Average Response Time';
    worksheet.mergeCells('B2:F2');
    worksheet.getCell('B2').value = 'Alerts';
    worksheet.getCell('B3').value = 'Critical'
    worksheet.getCell('C3').value = 'High'
    worksheet.getCell('D3').value = 'Medium'
    worksheet.getCell('E3').value = 'Low'
    worksheet.getCell('F3').value = 'Average Time'

    worksheet.mergeCells('G2:K2');
    worksheet.getCell('G2').value = 'Non Alerts';
    worksheet.getCell('G3').value = 'Critical'
    worksheet.getCell('H3').value = 'High'
    worksheet.getCell('I3').value = 'Medium'
    worksheet.getCell('J3').value = 'Low'
    worksheet.getCell('K3').value = 'Average Time'

    // Merge cells for "Resolution Time"
    worksheet.mergeCells('L1:U1');
    worksheet.getCell('L1').value = 'Average Resolution Time';
    worksheet.mergeCells('L2:P2');
    worksheet.getCell('L2').value = 'Alerts';
    worksheet.getCell('L3').value = 'Critical'
    worksheet.getCell('M3').value = 'High'
    worksheet.getCell('N3').value = 'Medium'
    worksheet.getCell('O3').value = 'Low'
    worksheet.getCell('P3').value = 'Average Time'

    worksheet.mergeCells('Q2:U2');
    worksheet.getCell('Q2').value = 'Non Alerts';
    worksheet.getCell('Q3').value = 'Critical'
    worksheet.getCell('R3').value = 'High'
    worksheet.getCell('S3').value = 'Medium'
    worksheet.getCell('T3').value = 'Low'
    worksheet.getCell('U3').value = 'Average Time'

    return { workbook, worksheet}
}
const { workbook, worksheet } = excelSheetInitialization()
    
// Calculating average Time spend by each client on each seperation
const calculateClientsAvgtimes = () => {

    let cellNameRow = 4

    clientsAllTimes.forEach((clientAllTimes)=>{

        let charCodeAscii = 65
        let cellNameCol = String.fromCharCode(charCodeAscii)
        const cellName = cellNameCol + cellNameRow
        worksheet.getCell(cellName).value = clientAllTimes['client']
        charCodeAscii += 1
        cellNameCol = String.fromCharCode(charCodeAscii)

        Object.keys(clientAllTimes).forEach(clientAllTimesKey => {
            if(clientAllTimesKey !== 'client'){
                Object.keys( clientAllTimes[clientAllTimesKey] ).forEach((alertType)=>{ 
                    // console.log("alertType",alertType)
                    let sumOfAveragePriorities = "0m"
                    let countOfAveragePrioritiesClient = 0
                    Object.keys(clientAllTimes[clientAllTimesKey][alertType]).forEach((priorityType)=>{
                        // console.log("priorityType",priorityType)

                        // Calculating averageTimeSpend of each seperation of client
                        const spendTimeOnPriorityString =  clientAllTimes[clientAllTimesKey][alertType][priorityType].spendTime
                        const {hours:spendTimeOnPriorityHours,minutes:spendTimeOnPriorityMinutes} = timeInHrsAndMins(spendTimeOnPriorityString)
                        const issueCountOfPriority = clientAllTimes[clientAllTimesKey][alertType][priorityType].issueCount
                        
                        if(spendTimeOnPriorityHours > 0 || spendTimeOnPriorityMinutes > 0 ){
                            const {avgTimeColonString:avgSpendTimeInHrs,avtTimeObjForm} = avgOfTime(spendTimeOnPriorityHours, spendTimeOnPriorityMinutes, issueCountOfPriority)
                            clientAllTimes[clientAllTimesKey][alertType][priorityType].avgSpendTime = avgSpendTimeInHrs

                            const avgTimeOfClientPriorityInHandM = `${avtTimeObjForm.hours}h ${avtTimeObjForm.minutes}m`

                            const {hours:averageTimeSectionHours,minutes:averageTimeSectionMinutes} = sumOfTime(sumOfAveragePriorities,avgTimeOfClientPriorityInHandM)
                            sumOfAveragePriorities = `${averageTimeSectionHours}h ${averageTimeSectionMinutes}m`
                            countOfAveragePrioritiesClient += 1
                            if(clientAllTimes['client'] === 'Comprinno' && clientAllTimesKey === 'avgResponseTime'&& alertType === 'alerts' ){
                                console.log("sumOfAveragePriorities,countOfAveragePrioritiesClient",sumOfAveragePriorities,countOfAveragePrioritiesClient)
                            }
                        }else{
                            clientAllTimes[clientAllTimesKey][alertType][priorityType].avgSpendTime = 'N/A'
                        }

                        const cellName = cellNameCol + cellNameRow
                        worksheet.getCell(cellName).value = clientAllTimes[clientAllTimesKey][alertType][priorityType].avgSpendTime
                        charCodeAscii += 1
                        cellNameCol = String.fromCharCode(charCodeAscii)
                        // Setting the Summary in clientAllTimes

                        if((spendTimeOnPriorityHours > 0 || spendTimeOnPriorityMinutes > 0) && clientAllTimes['client'] !== 'Summary'){
                            // console.log("summary",clientAllTimesKey,alertType,priorityType)
                            
                            const {avgTimeColonString:avgSpendTimeInHrs,avtTimeObjForm} = avgOfTime(spendTimeOnPriorityHours, spendTimeOnPriorityMinutes, issueCountOfPriority)
                            const timeSpendForSummary = avtTimeObjForm.hours + "h " + avtTimeObjForm.minutes + "m"
                            const spendTime = clientsAllTimes[clientsAllTimes.length -1][clientAllTimesKey][alertType][priorityType].spendTime
                            const {hours , minutes} = sumOfTime(spendTime,timeSpendForSummary)
                            clientsAllTimes[clientsAllTimes.length -1][clientAllTimesKey][alertType][priorityType].spendTime  = hours + "h " + minutes + "m"

                            // const timeSpendByClientInMillis = Math.floor(spendTimeOnPriority / issueCountOfPriority)
                    
                            // clientsAllTimes[clientsAllTimes.length -1][clientAllTimesKey][alertType][priorityType].spendTime += timeSpendByClientInMillis 
                            clientsAllTimes[clientsAllTimes.length -1][clientAllTimesKey][alertType][priorityType].issueCount += 1

                            // if(clientAllTimesKey === 'avgResponseTime'&& alertType === 'alerts' && priorityType ==="high"){
                            //     // const avgSpendTimeInMillis = Math.floor(spendTimeOnPriority / issueCountOfPriority)
                            //     // const avgSpendTimeInHrs = formattedTime(avgSpendTimeInMillis)
                            //     // console.log("Summary High",avgSpendTimeInHrs)                          
                            //     // const summaryTime = clientsAllTimes[clientsAllTimes.length -1][clientAllTimesKey][alertType][priorityType].spendTime 
                            //     // const summaryTimeInHrs = formattedTime(summaryTime)
                            //     // console.log("Summary High spendtimetotal",summaryTimeInHrs)
                            //     console.log("clients spendTime",spendTimeOnPriority)
                            //     console.log("sumary Highhhhhhhhhhhhhhh")
                            // }
                            // console.log("summary",JSON.stringify(clientsAllTimes[clientsAllTimes.length -1],null))
                        }

                    })

                    
                    const {hours:averageTimeClientHours,minutes:averageTimeClientMinutes} = timeInHrsAndMins(sumOfAveragePriorities)
                    // console.log("outside",averageTimeClientHours,"minutes",averageTimeClientMinutes,"count",countOfAveragePrioritiesClient)
                    let averageTime;
                    if(countOfAveragePrioritiesClient > 0){
                        const { avgTimeColonString , avtTimeObjForm } = avgOfTime(averageTimeClientHours, averageTimeClientMinutes, countOfAveragePrioritiesClient)
                        averageTime = avgTimeColonString
                    }else{
                        averageTime = 'N/A'
                    }

                    const cellName = cellNameCol + cellNameRow
                    worksheet.getCell(cellName).value = averageTime
                    charCodeAscii += 1
                    cellNameCol = String.fromCharCode(charCodeAscii)
                })
            }
        })
        cellNameRow += 1
    })
}
calculateClientsAvgtimes()



const excelSheetCreating = async () => {
    const fileName = (new Date).getTime() + 'average-time-automation.xlsx'
    await workbook.xlsx.writeFile(fileName)
}
excelSheetCreating()


// Open representation of object
// const jsonArray = clientsAllTimes.map(obj => JSON.stringify(obj,null,2))
// jsonArray.forEach(json => console.log(json))






























