
const express = require('express');
const axios   = require('axios');
const qs = require('qs');
const fillExcelTimesheet = require('./timesheet.js');
var bodyParser = require('body-parser');
var path = require("path");


module.exports = getMicrosoftToken
const app  = express();
const port = 3000;


app.use(bodyParser.urlencoded({extended: false}));
app.use(express.static(path.join(__dirname,'./public')));

app.post('/download', (req, res) => {
    const filePath = '/tmp/file.xlsx'; // Replace with the actual file path
  
    console.log(filePath)
    res.download(filePath, 'timesheet.xlsx', (err) => {
      if (err) {
        console.error(err);
        res.status(500).send('Error downloading the file');
      }
    });
  });
  
  // Define a route to navigate to another HTML file
  app.get('/success', (req, res) => {
    res.sendFile('success.html', { root: __dirname });
  });
  app.get('/failure', (req, res) => {
    res.sendFile('failure.html', { root: __dirname });
  });

  
app.get('/home', function (req, res) {
    res.sendFile('index.html', { root: __dirname });
});

Array.prototype.last = function() {
    return this[this.length - 1];
}

Array.prototype.first = function() {
    return this[0];
}


  
var response = null
var sendEmail = false
var isUAE = true
app.post('/submitTimesheet', (req, res) => {
    console.log(req.body)
    
    const email = req.body.email;
    const password = req.body.password;
    sendEmail = req.body.sendEmail;
    isUAE = req.body.isUAE;
    
    if (!email || !password || !email.includes('systemsltd')) {
        return res.status(400).json({ error: 'Wrong email or password!' });
    }
    response = res;
    getMicrosoftToken(email, password);
});


async function getMicrosoftToken(email, pass) {

    const data = qs.stringify({
        client_id: 'f4879a42-d86f-45b2-b130-293d4aedf10c',
            scope: 'https://graph.microsoft.com/.default',
            redirect_uri: 'https://syshcm.systemsltd.com/EssPlus/login',
            grant_type: 'password',
            username: email,
            password: pass,
      });

    axios.post('https://login.microsoftonline.com/def44f5f-0783-4b05-8f2f-dd615c5dfec4/oauth2/v2.0/token', 
    data, {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
    }).then(res => {
            const accessToken = res.data.access_token;
            getSystemsAccessToken(accessToken, email);
    })
    .catch(error => {
        response.status(400).json({ error: 'Please check the email and password are correct!' });
    });
}

async function getSystemsAccessToken(accessToken, email) {

    const data = qs.stringify({
        username: email,
        password: 'undefined',
        grant_type: 'password',
        scope: 'openid email phone profile offline_access roles syshcm_api',
        client_id: 'syshcm_ess',
        azure_access_token: accessToken,
      });

    axios.post('https://syshcm.visionetsystems.com/syshcmapi/connect/token', 
    data, {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
    }).then(res => {
            const accessToken = res.data.access_token
            // submitExcelsheet(accessToken);
            getTimesheetList(accessToken);
    })
    .catch(error => {
        response.status(400).json({ error: 'Please check the email and password are correct!' });
    });
}

async function getTimesheetList(accessToken) {
    axios.get('https://syshcm.visionetsystems.com/syshcmapi/api/v1/Timesheet/Details?pageNumber=1&pageSize=10&sortOrder=&Status=&ProjectId=&TaskId=&StartDate=&EndDate=', {
        headers: {
            'authorization': `Bearer ${accessToken}`,
          },
    }).then(res => {
        
        var customerId = '0'
        var projectId = '0'
        var roleId = '0'
        var accountId = '0'
        var billId = '0'
        var taskId = '0'

        const isTimesheetCompleted = res.data.filter((function(element){
            return (element.status == "Missing")
        })).length == 0

        if(isTimesheetCompleted)
            response.redirect('/failure')


        var list = res.data.filter((function(element){
            return (element.projectInfos.length > 0)
        }))
        
        if(list.length > 0) {
            customerId = list[0].projectInfos[0].customer.value;
            projectId = list[0].projectInfos[0].project.value;
            roleId = list[0].projectInfos[0].roles.value;
            accountId = list[0].projectInfos[0].account.value;
            billId = list[0].projectInfos[0].accrualStatusId;
            taskId = list[0].projectInfos[0].task.value;
        }

        res.data.forEach((function(element) 
        { 
            const status = element.status;
            if(status == "Missing"){
            const empId = element.holidays[0].empId;
            
            const dates = element.holidays.map((function(element){
                return element.date.split('T')[0];
            }))
            
            const lastDay = element.weekDate.split('T')[0];
            const timesheetRequestId = element.timesheetRequestId;

            const requestArray = {
                "timesheetRequestDetails": [],
                "employeeTimesheetRequestId": 0,
                "timesheetStatusId": 0,
                "weekEndDate": lastDay,
                "approverId": null,
                "activeInd": true,
                "approverComment": "",
                "empId": empId,
                "employeeComment": " "
            }
            
            for(let i = 0; i < 7; i++){
                var hours = 8
                if(i == 0 || i == dates.length - 1) {
                    hours = null
                }
                
                const obj = {
                    "employeeTimesheetRequestId": timesheetRequestId,
                    "employeeTimesheetRequestDetailsId": 0,
                    "hours": hours,
                    "accrualStatusId": null,
                    "date": dates[i],
                    "projectId": projectId,
                    "billId": billId,
                    "taskId": taskId,
                    "comment": "",
                    "customerId": customerId,
                    "roleId": roleId,
                    "accountId": accountId
                }

                requestArray.timesheetRequestDetails.push(obj);
              }

              submitMissingTimesheet(requestArray, accessToken);
              
            }
        }));
        
    })
    .catch(error => {
        console.error(error);
        response.status(400).json({ error: 'An error occurred with retrieving the timesheet!' });
    });
}


async function submitMissingTimesheet(requestArray, token) {
    axios.post('https://syshcm.visionetsystems.com/syshcmapi/api/v1/Timesheet/submit', 
    requestArray, {
        headers: {
            'authorization': `Bearer ${token}`,
          },
    }).then(res => {
        if(res.status == 200){
            if(sendEmail)
                submitExcelsheet(token);
            else
                response.status(200).json({ succeess: 'Timesheet has been submited ya abo elso7ab!' });
            }
        else
            response.status(400).json({ error: 'An error occurred!' });
        
    })
    .catch(error => {
        response.status(400).json({ error: 'An error occurred with submitting the timesheet!' });
    });
}





async function submitExcelsheet(accessToken) {
    axios.get('https://syshcm.visionetsystems.com/syshcmapi/api/v1/Timesheet/Details?pageNumber=1&pageSize=10&sortOrder=&Status=&ProjectId=&TaskId=&StartDate=&EndDate=', {
        headers: {
            'authorization': `Bearer ${accessToken}`,
          },
    }).then(res => {
        
        const map = {};
        const splittedDate = res.data[0].weekDate.split('-');
        var thisMonth = `${splittedDate[0]}-${splittedDate[1]}`;
        let myDate = new Date()
        let myMonth = `${myDate.getFullYear()}-${myDate.getMonth() + 1}`
        console.log(myMonth)
        // thisMonth = '2023-10'
        
        res.data.forEach((function(element) {
            if(element.status != "Missing"){

                const hours = element.projectInfos[0].hours.map((function(item){
                    return item.hour
                }));
    
                const dates = element.holidays.map((function(item){
                    return item.date.split('T')[0];
                }));
    
                
                for (let i = 0; i < dates.length; i++) {
                    if(dates[i].includes(thisMonth))
                        map[dates[i]] = hours[i];
                }

            }
        }));
        
        let sorted_obj = Object.keys(map)
                .sort().reduce((temp_obj, key) => {
                temp_obj[key] = map[key];
                return temp_obj;
                }, {});   

        let date = `${getMonthName(thisMonth.split('-')[1]).substring(0, 3)}-${thisMonth.split('-')[0].substring(2,4)}`
        let lastArray = Array.from(Object.values(sorted_obj));
        let datesList = Array.from(Object.keys(sorted_obj));
        let location = 'KSA' 
        if(isUAE) 
            location = 'UAE'
        fillExcelTimesheet(response, lastArray, date, datesList, location);

        
    }).catch(error => {
        console.error(error);
        response.status(400).json({ error: 'An error occurred with retrieving the timesheet!' });
    });
}


function getMonthName(monthNumber) {
    const date = new Date();
    date.setMonth(monthNumber - 1);
  
    return date.toLocaleString('en-US', {
      month: 'long',
    });
  }


app.listen(port, () => {

});