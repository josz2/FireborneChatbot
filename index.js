'use strict';

const functions = require('firebase-functions');
const {google} = require('googleapis');
const {WebhookClient} = require('dialogflow-fulfillment');
const BIGQUERY = require('@google-cloud/bigquery');


// Enter your calendar ID below and service account JSON below
const calendarId = "c_i1oimq9po0tjvk1vnqoq8ptpi8@group.calendar.google.com";
const serviceAccount = {"type": "service_account",
  "project_id": "chatbot-appointment-scheduler",
  "private_key_id": "7f138dc98a089671b04655a64740cfc01b884764",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDG6XdWPCKLHaHS\nJ+KIh66BcCyjEH7PF5LJqOHVe2g9tfs3XquTaUHYmI6pZPADdLqo07z4vj75vKWw\nV9UFDd8p07F/WjjXXWh0DOrRW1yjfEX7We60La1vjtPIQ7stykDD1jrpA143Re4X\nI1zzTyjpAgVszAwXp8IqqR0y8NiSQgk8kwhme0gRLHIkLIbvhWBbvywPd6UVubqf\nrL2anNg+fNWRgmqlZvo5YRGaJiLSgZETj21QgRammm1ydSEGRDo0iic2qHcg/TCq\n4yqr0cPmv3EnZAjv+9Ij6CfKoqMkSmLEIL+CNBTLDdms7K6dXs3HHA9Bt6Av8ejq\nfPuOkYrBAgMBAAECggEAF8XbJwWuoLcIyjHNYaEqtmpaUV5VM4nHF74tCHtUbuyv\nIrPYOHjf32PDSCRJybXzoZ4Vw5pUMzuMR2ZF7bHV2j1EZVq6eKXfqCALl/wt9xWR\nFRxvWp4rcG3u6oxKxIsbwLQbzBHEmsFLNn153FP5iRieXp2H8/NPMdNOq7IMhjUz\nbX/fwYhsATQqtfOaeno8LmsPN7XB7nN5Y/zsGoSa8ItFed2fm0m4+frk54T2qtWL\nDFkx1BBXeei8YfAAcSgu9chl6jF/2arMl69Qk8KjdUgFcmSUxWNoLSiMs7iHiSjj\nCXyo5AeytjQGwA7jgZepuERUD+iOLF4lVRS8eKczWQKBgQDxtYrZe/ung3M/etFw\nUYXAp2XCrLgiOSk61EgzYVwTsSChYdgEk4wFX12o2ea1pRViCs5UT8YiA44IVrpp\n1KljlxbB7Yr1MRexFyrDLKJo6lKIDZ90Xe2ypP+MgRp1oXKzjlHXbZZtjpe/6Lo8\nkDUhvZ1gd78LzfnVVgM7t8jcyQKBgQDSrCetkqi5z1vyFBx9J+i3KZDiR0Yp8UsM\nFTLqLVU6rtRk0LILJXNjoHL4kQ6BRbdT4GC1FHtqv3wMtZR0eQbwttOCF+iULF1j\n8YJxY67EHsXPq98r465l8YxANHjP5PhSKNBm2clCAFX1KGOI1Byb1RGGQeNdewG1\nJBDImyBSOQKBgArQSGn6dgPEib9pSz1vKEC6PH89Iu/FBucu4BwMWwY2gnM14Wgz\nAayr25DWTtAJlq9QNHLpLsAO0Kfm2Wgqr3lZJRd//RuDGsA9fRhGQu3WreKQWXXn\nTd8UKqqqi/h/RJZr45VzvashGgDn9I0JFpdv2D6cnNt2V5sHwhVF36KhAoGBAKp4\nodbDMQLB9y3I9lCUBayI1vMzJ2RzGv4Y/U0fB7NnmvhFI3z/fgKk58OZZTpX1oPp\nsXd1rnRvpAqIuCsTb/lCh53iiNG1oJBp8dqdBeMu33QvKHRUVV+qeInPq97V8dZR\nrmk7W66rpOKvHvOuZ8P1Qqv4DuoqyfPwzh/13s6JAoGAcddwpnKyL8/378KtHE3t\nSL85ZrQbdSKcW2RPmCc+QChjQB32CbaRNH/ovVHtHBTPbl1jWII6liHx0dpYromd\nsNgqOlMdJMU+zAzN7e2ON7v4zDVjer7pBrqLVHRpmvXeMm3kT0MD6JxkU5C/xCu/\nQGu7rkm2lpHbLPA2dKRbv5Q=\n-----END PRIVATE KEY-----\n",
  "client_email": "appointment-scheduler@chatbot-appointment-scheduler.iam.gserviceaccount.com",
  "client_id": "106662332566927876858",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/appointment-scheduler%40chatbot-appointment-scheduler.iam.gserviceaccount.com"
}; // Starts with {"type": "service_account",...

// Set up Google Calendar Service account credentials
const serviceAccountAuth = new google.auth.JWT({
  email: serviceAccount.client_email,
  key: serviceAccount.private_key,
  scopes: 'https://www.googleapis.com/auth/calendar'
});

const calendar = google.calendar('v3');
process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements

const timeZone = 'America/New_York';
const timeZoneOffset = '-05:00';

exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
  const agent = new WebhookClient({ request, response });
  console.log("Parameters", agent.parameters);
  const appointment_type = agent.parameters.AppointmentType;

// Function to create appointment in calendar  
function makeAppointment (agent) {
    // Calculate appointment start and end datetimes (end = +1hr from start)
    const dateTimeStart = new Date(Date.parse(agent.parameters.date.split('T')[0] + 'T' + agent.parameters.time.split('T')[1].split('-')[0] + timeZoneOffset));
    const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
    const appointmentTimeString = dateTimeStart.toLocaleString(
      'en-US',
      { month: 'long', day: 'numeric', hour: 'numeric', timeZone: timeZone }
    );
  
// Check the availability of the time, and make an appointment if there is time on the calendar
    return createCalendarEvent(dateTimeStart, dateTimeEnd, appointment_type).then(() => {
      agent.add(`Ok, let me see if we can fit you in. ${appointmentTimeString} is fine!.`);

// Insert data into a table
      addToBigQuery(agent, appointment_type);
    }).catch(() => {
      agent.add(`I'm sorry, there are no slots available for ${appointmentTimeString}.`);
    });
  }

  let intentMap = new Map();
  intentMap.set('Schedule Appointment', makeAppointment);
  agent.handleRequest(intentMap);
});

//Add data to BigQuery
function addToBigQuery(agent, appointment_type) {
    const date_bq = agent.parameters.date.split('T')[0];
    const time_bq = agent.parameters.time.split('T')[1].split('-')[0];
    /**
    * TODO(developer): Uncomment the following lines before running the sample.
    */
    const projectId = '<chatbot-appointment-scheduler>'; 
    const datasetId = "<Appointment_dataset>";
    const tableId = "<Appointment_table>";
    const bigquery = new BIGQUERY({
      projectId: projectId
    });
   const rows = [{date: date_bq, time: time_bq, type: appointment_type}];
  
   bigquery
  .dataset(datasetId)
  .table(tableId)
  .insert(rows)
  .then(() => {
    console.log(`Inserted ${rows.length} rows`);
  })
  .catch(err => {
    if (err && err.name === 'PartialFailureError') {
      if (err.errors && err.errors.length > 0) {
        console.log('Insert errors:');
        err.errors.forEach(err => console.error(err));
      }
    } else {
      console.error('ERROR:', err);
    }
  });
  agent.add(`Added ${date_bq} and ${time_bq} into the table`);
}

// Function to create appointment in google calendar  
function createCalendarEvent (dateTimeStart, dateTimeEnd, appointment_type) {
  return new Promise((resolve, reject) => {
    calendar.events.list({
      auth: serviceAccountAuth, // List events for time period
      calendarId: calendarId,
      timeMin: dateTimeStart.toISOString(),
      timeMax: dateTimeEnd.toISOString()
    }, (err, calendarResponse) => {
      // Check if there is a event already on the Calendar
      if (err || calendarResponse.data.items.length > 0) {
        reject(err || new Error('Requested time conflicts with another appointment'));
      } else {
        // Create event for the requested time period
        calendar.events.insert({ auth: serviceAccountAuth,
          calendarId: calendarId,
          resource: {summary: appointment_type +' Appointment', description: appointment_type,
            start: {dateTime: dateTimeStart},
            end: {dateTime: dateTimeEnd}}
        }, (err, event) => {
          err ? reject(err) : resolve(event);
        }
        );
      }
    });
  });
}