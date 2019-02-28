const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const util = require('util')

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/calendar.readonly', 'https://www.googleapis.com/auth/calendar.events'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

const all_floor = [
  {generatedResourceName: "T02_3F会議室-【3T-01】(4名)", resourceEmail: "dentsudigital.co.jp_3737343736393935383334@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-02】(4名)", resourceEmail: "dentsudigital.co.jp_3632333635303132353537@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-03】(4名)", resourceEmail: "dentsudigital.co.jp_3837393933333439323639@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-04】(10名)", resourceEmail: "dentsudigital.co.jp_3631383137393337333234@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-05】(8名)", resourceEmail: "dentsudigital.co.jp_3332363331383634323039@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-06】(8名)", resourceEmail: "dentsudigital.co.jp_3137353438353234353832@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-07】(8名)", resourceEmail: "dentsudigital.co.jp_32343632363433343035@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-08】(18名)", resourceEmail: "dentsudigital.co.jp_3930393936313137333936@resource.calendar.google.com"},
  {generatedResourceName: "T02_3F会議室-【3T-09】(12名)", resourceEmail: "dentsudigital.co.jp_363936363138323935@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-01】(10名)", resourceEmail: "dentsudigital.co.jp_3633313937313138353635@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-02】(10名)", resourceEmail: "dentsudigital.co.jp_35303937363138373334@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-03】(10名)", resourceEmail: "dentsudigital.co.jp_3733313938313138333436@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-04】(10名)", resourceEmail: "dentsudigital.co.jp_38313939303038333738@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-05】(5名)", resourceEmail: "dentsudigital.co.jp_31313939343734383833@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-06】(8名)", resourceEmail: "dentsudigital.co.jp_343539393937343833@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-07】(12名)", resourceEmail: "dentsudigital.co.jp_35313238393238373933@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5G-08】(8名)", resourceEmail: "dentsudigital.co.jp_333437393734323335@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-01】(4名・ソファ席)", resourceEmail: "dentsudigital.co.jp_393331383436343137@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-02】(9名)", resourceEmail: "dentsudigital.co.jp_31333732333436313834@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-03】(9名)", resourceEmail: "dentsudigital.co.jp_39383632383436353339@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-04】(12名・個別机)", resourceEmail: "dentsudigital.co.jp_3832343837393735333830@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-05】(20名・タブレットチェア)", resourceEmail: "dentsudigital.co.jp_33313933383436363634@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-06】(8名・昇降デスクチェア)", resourceEmail: "dentsudigital.co.jp_3730313232363536393236@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-07】(12名)", resourceEmail: "dentsudigital.co.jp_38313235313137363139@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-【5T-08】(12名)", resourceEmail: "dentsudigital.co.jp_38303335363137313637@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-プレゼンルーム (20名)", resourceEmail: "dentsudigital.co.jp_33313836313137343631@resource.calendar.google.com"},
  {generatedResourceName: "T02_5F会議室-応接室 (4名)", resourceEmail: "dentsudigital.co.jp_33303436363137353130@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-01】(4名)", resourceEmail: "dentsudigital.co.jp_36323137333838373835@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-02】(4名)", resourceEmail: "dentsudigital.co.jp_34373037383838393335@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-03】(4名)", resourceEmail: "dentsudigital.co.jp_33313938333838363632@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-04】(4名)", resourceEmail: "dentsudigital.co.jp_31363838383838373833@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-05】(4名)", resourceEmail: "dentsudigital.co.jp_313739333838353431@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-06】(9名)", resourceEmail: "dentsudigital.co.jp_36313130323738343938@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-07】(12名)", resourceEmail: "dentsudigital.co.jp_3434343130373738333135@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-08】(9名)", resourceEmail: "dentsudigital.co.jp_3735393131323738343436@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-09】(9名)", resourceEmail: "dentsudigital.co.jp_3630383131373738383332@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-10】(12名)", resourceEmail: "dentsudigital.co.jp_31323338353335323138@resource.calendar.google.com"},
  {generatedResourceName: "T02_9F会議室-【9T-11】(9名)", resourceEmail: "dentsudigital.co.jp_3133353133313232393535@resource.calendar.google.com"}];
  
  var result_all_floor = [];
// Load client secrets from a local file.
    const run = function(data, cp){
       return new Promise(function(resolve, reject){
        fs.readFile('credentials.json', (err, content) => {
          if (err) 
              return console.log('Error loading client secret file:', err);
        // Authorize a client with credentials, then call the Google Calendar API.
        resolve(authorize(JSON.parse(content), listEvents, data, cp))
       })
    

});}

const createEventGsuite = function (data){
  return new Promise(function(resolve, reject){
    fs.readFile('credentials.json', (err, content) => {
      if (err) 
          return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Calendar API.
    resolve(authorize(JSON.parse(content), createEvent, data))
   })


});
}

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback, data, cp) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  return new Promise(function(resolve, reject){
    fs.readFile(TOKEN_PATH, (err, token) => {
      if (err) return getAccessToken(oAuth2Client, callback);
      oAuth2Client.setCredentials(JSON.parse(token));
      resolve(callback(oAuth2Client, data, cp))
    });
  })
  
  //return "FFFFF"
}

function getAccessToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

function listEvents(auth, data, cp) {
  return func(auth, data, cp);

}

function createEvent(auth, data){
  return func2(auth, data);
}

async function func (auth, data, cp){
  result_all_floor = []
  this.au = auth
  var times = data
  var capacity = cp;
  var count = 0;

  // console.log("+++++++++ ", data)


      var all_floor_temp = []; 
        
            // console.log("****** ", times[i])
            all_floor_temp= all_floor.slice();

            while (result_all_floor.length<3 && all_floor_temp.length>0)
            {
                var random = Math.floor(Math.random()*all_floor_temp.length);
                var item = all_floor_temp[random];
                all_floor_temp.splice(random, 1);
                // console.log(">>>>>> 3F ",item.generatedResourceName)
                for (var i = 0; i<times.length; i++)
                {
                count++;
                if(enoughCapacity(item.generatedResourceName, capacity)) {
                        for (var i = 0; i<times.length; i++)
                        {
                          let data = {'time': times[i].time, 
                            'location': [
                                  { 
                                    displayName: item.generatedResourceName,
                                    locationEmailAddress: item.resourceEmail
                                  }]}
                          // console.log("****** ", times[i])
                          let status = await checkRoomAvailability(item, times[i]);
                          // console.log("############# Result ",status)
                          if(status){
                            result_all_floor.push(data)
                          }
                    
                    }
                }
              }
            }

        // console.log(util.inspect(result_all_floor, {showHidden: false, depth: null}))
        return {data : result_all_floor, count:count};



}


async function func2 (auth, data){
  var calendar = google.calendar({version: 'v3', auth:this.au});
  calendar.events.insert({
    auth: auth,
    calendarId: 'primary',
    resource: data,
  }, function(err, event) {
    if (err) {
      console.log('There was an error contacting the Calendar service: ' + err);
      return false;
    }
    console.log('Event created: %s', event);
    return true;
  });
}





function checkRoomAvailability (room, time)
  {
        var date = new Date(time.time.start);
        var calendar = google.calendar({version: 'v3', auth:this.au});
        var convertedDate = date.getFullYear()+'-'+(date.getMonth()+1)+'-'+date.getDate();

      return new Promise(function(resolve, reject) {
        calendar.events.list({
          calendarId: room.resourceEmail,
          timeMin: (new Date(Date.parse(convertedDate+" 00:00:00"))).toISOString(),
          timeMax: (new Date(Date.parse(convertedDate+" 24:00:00"))).toISOString(),
          singleEvents: true,
          showDeleted:false,
          orderBy: 'startTime',
        }

        , (err, res) => {

            if (res) {
                  const events = res.data.items;
                  resolve(haveSpace(time, events))
            }

          });
      });

  }

  function haveSpace(time, events)
  {
    var temp = true;

    if(events!=undefined){
        for(var event of events){
            var date2 = {
              start: event.start.dateTime || event.start.date,
              end: event.end.dateTime || event.end.date
            }
            date2.start = date2.start.substring(0, date2.start.length-6);
            date2.end = date2.end.substring(0, date2.end.length-6);
    
           if ((new Date(time.time.start) < new Date(date2.end))&&(new Date(date2.start) < new Date(time.time.end))){
              temp = false;
              break;
           }
                
          }
    }
    else
        temp = false
    

    return temp;
  }

  function enoughCapacity(room, capacity)
  {
    var regExp = /\(([^)]+)\)/;
    var matches = regExp.exec(room);
    var temp = matches[1];
    temp = temp.substring(0, temp.length - 1)
    var number = parseInt(temp);
    if(capacity<number)
      return true;

    return false;
  }

  module.exports = {
      getGsuiteRooms: run,
      createEventGsuite: createEventGsuite
  } 