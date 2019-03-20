var graph = require('@microsoft/microsoft-graph-client');
var tools = require('./gsuite');
const axios = require('axios');
const util = require('util');
const moment = require('moment');

module.exports = {
    getUserDetails: async function(accessToken) {
        const client = getAuthenticatedClient(accessToken);

        const user = await client.api('/me').get();
        return user;
    },


    user_list: async function(accessToken) {
        var  client = getAuthenticatedClient(accessToken);
        // console.log(accessToken)
        // var nextLink;
        const users = await client
            .api('/me/people')
            .select("id","displayName","scoredEmailAddresses")
            .top(15)
            .get();
        users.value.forEach((user)=>{
            user['mail'] = user.scoredEmailAddresses[0].address;
            delete user['scoredEmailAddresses'];
        })
        return users.value;
    },

    searchUser: async function(accessToken, data) {
        var email, name;
        if(data.includes("@"))
            email = data;
        else
            name = data;
        var  client = getAuthenticatedClient(accessToken);
        // console.log(accessToken);
        var filter;

        if(email==""||email==null)
            filter = "startswith(displayName, '"+name+"')";
        else
            filter = "startswith(mail, '"+email+"')";

        const user = await client
            .api('/users/')
            .select("id","displayName","mail")
            .filter(filter)
            .get();
        var obj ={};
        if (user.value.length!=1) {
            obj = { 'status': false };
        }
        else
            obj = { 'status': true, 'data': user.value[0] };
        return obj;
    },

    findMeetingTimes: async function(accessToken, info) {
        console.log("In findMeetingTimes")
        var res =[]
        var headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken,
            'Prefer': 'outlook.timezone="Asia/Tokyo"'
        }

        data = {
            "attendees": [],
            'maxCandidates': 3,
            "locationConstraint": {
                "isRequired": "true",
                "suggestLocation": "false",
                "locations": [
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_London",
                        "locationEmailAddress": "JPTYO6.RM0010@globalloc.com"
                    },
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Chicago",
                        "locationEmailAddress": "JPTYO6.RM0011@globalloc.com"
                    },
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Singapore",
                        "locationEmailAddress": "JPTYO6.RM0013@globalloc.com"
                    },

                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Shanghai",
                        "locationEmailAddress": "JPTYO6.RM0012@globalloc.com"
                    },


                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_SaoPaulo",
                        "locationEmailAddress": "JPTYO6.RM0014@globalloc.com"
                    }

                ]
            },
            "meetingDuration": "",
            "returnSuggestionReasons": "true",
            "minimumAttendeePercentage": "100"
        }

        info['attendees'].forEach((attendee)=>{
            var obj = {
                "type": "required",
                "emailAddress": {
                    "address": attendee
                }
            }
            data['attendees'].push(obj)
        })

        data['meetingDuration'] = moment.duration(info['duration'], 'minutes').toISOString()

        console.log("DDDDD ", data);

        await axios.post('https://graph.microsoft.com/v1.0/me/findMeetingTimes', data, {headers: headers})

            .then((response) => {
                response.data['meetingTimeSuggestions'].forEach((e)=>{
                    var result = {
                        time:{
                            start:e['meetingTimeSlot']['start']['dateTime'],
                            end:e['meetingTimeSlot']['end']['dateTime']
                        },

                        locations:e['locations']

                    }
                    res.push(result)

                })
                console.log(res);
                console.log("******")
            })
            .catch((error) => {
                console.log("Response ",error)
            })
        const result = await tools.getGsuiteRooms(res, data['attendees'].length)
        return {'internal':res,
                'external':result};
    },

    forceMeetingTimes: async function(accessToken, info) {
        var result ;
        var headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer '+accessToken,
            'Prefer' : 'outlook.timezone="Asia/Tokyo"'
        }

        data = {
            "locationConstraint": {
                "isRequired": "true",
                "suggestLocation": "false",
                "locations": [
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_London",
                        "locationEmailAddress": "JPTYO6.RM0010@globalloc.com"
                    },
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Chicago",
                        "locationEmailAddress": "JPTYO6.RM0011@globalloc.com"
                    },
                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Singapore",
                        "locationEmailAddress": "JPTYO6.RM0013@globalloc.com"
                    },

                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_Shanghai",
                        "locationEmailAddress": "JPTYO6.RM0012@globalloc.com"
                    },


                    {
                        "resolveAvailability": "true",
                        "displayName": "JPTYO6 - Room_SaoPaulo",
                        "locationEmailAddress": "JPTYO6.RM0014@globalloc.com"
                    }

                ]
            },
            "timeConstraint": {
                "activityDomain":"unrestricted",
                "timeslots": [
                    {
                        "start" : {
                            "dateTime": "",
                            "timeZone": "Tokyo Standard Time"
                        },
                        "end" : {
                            "dateTime": "",
                            "timeZone": "Tokyo Standard Time"
                        }
                    }
                ]
            },
            "returnSuggestionReasons": "true",
            "minimumAttendeePercentage": "0",
            "meetingDuration":""
        }

        data['meetingDuration'] = moment.duration(info['duration'], 'minutes').toISOString()
        data['timeConstraint']['timeslots'][0]['start']['dateTime'] = info['date']+'T'+info['time']+'Z'

        var hour = parseInt(info['time'].split(":")[0]) + time_convert(info['duration']).hours
        var minute = parseInt(info['time'].split(":")[1]) + time_convert(info['duration']).minutes

        if(minute>=60){
            hour+=1;
            minute-=60;
        }

        data['timeConstraint']['timeslots'][0]['end']['dateTime'] = info['date']+'T'+pad(hour,2)+':'+pad(minute,2)+":00"+'Z'

        // data['timeConstraint']['timeslots'][0]['end']['dateTime'] = '2019-03-08T09:30:00Z'

        console.log(">>>>>>><<<<<<<<<< ",info)
        console.log(">>>>>>><<<<<<<<<< ",data.timeConstraint.timeslots)
        console.log(">>>>>>><<<<<<<<<< ",data)


        await axios.post('https://graph.microsoft.com/v1.0/me/findMeetingTimes', data, {headers: headers})

            .then((response) => {
                var temp = response.data['meetingTimeSuggestions'][0]
                console.log(">>>>>>><<<<<<<<<< ",response)
                result = {
                    status:false,
                    time:{
                        start:temp['meetingTimeSlot']['start']['dateTime'],
                        end:temp['meetingTimeSlot']['end']['dateTime']
                    },

                    locations:temp['locations']

                }
                if (result['locations'].length>0)
                    result['status'] = true

            })
            .catch((error) => {
                console.log("Response ",error)
            })
        return result;
    },

    createEvent: async function(accessToken, info) {
        console.log("GGGGGGGGGGGG")
        console.log(util.inspect(info, {showHidden: false, depth: null}))
        const res = []
        var result, status = false;
        var headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer '+accessToken
        };
        var event = {
            "start": {
                "dateTime": "",
                "timeZone": "Tokyo Standard Time"
            },
            "end": {
                "dateTime": "",
                "timeZone": "Tokyo Standard Time"
            },
            "attendees": [{

            }]
        }
        event.subject = info.subject
        event.start.dateTime = info.start
        event.end.dateTime = info.end
        event.location = info.location
        event.attendees = info.attendees
        console.log("EEEEEE "+event)
        await axios.post('https://graph.microsoft.com/v1.0/me/calendar/events', event, {headers: headers})

            .then((response) => {
                console.log("**********************")
                console.log(response)
                status =true;
            })
            .catch((error) => {
                console.log("Response ",error)
                status = false;
            })
        if(info.category.includes("External") || info.category.includes("external")|| info.category.includes("外部"))
            {
                var userr = await this.getUserDetails(accessToken)
                console.log("######## $$$$$$$$$ ",userr.displayName)
                var event_google = {
                    'summary': userr.displayName+"_"+info.subject,
                    // 'location': info.location.displayName,
                    'start': {
                    'dateTime': info.start.slice(0,19)+'+09:00',
                    'timeZone': 'Asia/Tokyo',
                    },
                    'end': {
                    'dateTime': info.end.slice(0,19)+'+09:00',
                    'timeZone': 'Asia/Tokyo',
                    },
                    'attendees': [
                    {
                        'email': info.location.locationEmailAddress,
                        'resource': true
                    },
                    {
                        'email': 'dentsu-isobar-01@cci.co.jp',
                        'responseStatus':'accepted'
                    }
                    ]
                }
                await tools.createEventGsuite(event_google)
            }
        return status;
    },
};

function getAuthenticatedClient(accessToken) {
    // Initialize Graph client
    const client = graph.Client.init({
        // Use the provided access token to authenticate
        // requests
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    return client;
}

function time_convert(num)
 {
  var hours = Math.floor(num / 60);
  var minutes = num % 60;
  return {hours, minutes};
}

function pad(num, size) {
    var s = num+"";
    while (s.length < size) s = "0" + s;
    return s;
}

