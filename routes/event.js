// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
var mailto;
var mailsubject;
var mailbody;
var personName;
var roomadd;
var roomname;
var starttime;
var endtime;
//var bobmsg;
function sortProperties(obj) {
    // convert object into array
    var sortable = [];

    for (var key in obj)
        if (obj.hasOwnProperty(key))
            sortable.push([key, obj[key]]); // each item is an array in format [key, value]

    // sort items by value
    sortable.sort(function (a, b) {
        var x = a[1],
            y = b[1];
        return x < y ? -1 : x > y ? 1 : 0;
    });
    return sortable; // array in format [ [ key1, val1 ], [ key2, val2 ], ... ]
}


/* GET /contacts */
router.get('/', async function (req, res, next) {

    var dateFormat = function () {
        var token = /d{1,4}|m{1,4}|yy(?:yy)?|([HhMsTt])\1?|[LloSZ]|"[^"]*"|'[^']*'/g,
            timezone =
            /\b(?:[PMCEA][SDP]T|(?:Pacific|Mountain|Central|Eastern|Atlantic) (?:Standard|Daylight|Prevailing) Time|(?:GMT|UTC)(?:[-+]\d{4})?)\b/g,
            timezoneClip = /[^-+\dA-Z]/g,
            pad = function (val, len) {
                val = String(val);
                len = len || 2;
                while (val.length < len) val = "0" + val;
                return val;
            };

        // Regexes and supporting functions are cached through closure
        return function (date, mask, utc) {
            var dF = dateFormat;

            // You can't provide utc if you skip other args (use the "UTC:" mask prefix)
            if (arguments.length == 1 && Object.prototype.toString.call(date) == "[object String]" && !/\d/
                .test(date)) {
                mask = date;
                date = undefined;
            }

            // Passing date through Date applies Date.parse, if necessary
            date = date ? new Date(date) : new Date;
            if (isNaN(date)) throw SyntaxError("invalid date");

            mask = String(dF.masks[mask] || mask || dF.masks["default"]);

            // Allow setting the utc argument via the mask
            if (mask.slice(0, 4) == "UTC:") {
                mask = mask.slice(4);
                utc = true;
            }

            var _ = utc ? "getUTC" : "get",
                d = date[_ + "Date"](),
                D = date[_ + "Day"](),
                m = date[_ + "Month"](),
                y = date[_ + "FullYear"](),
                H = date[_ + "Hours"](),
                M = date[_ + "Minutes"](),
                s = date[_ + "Seconds"](),
                L = date[_ + "Milliseconds"](),
                o = utc ? 0 : date.getTimezoneOffset(),
                flags = {
                    d: d,
                    dd: pad(d),
                    ddd: dF.i18n.dayNames[D],
                    dddd: dF.i18n.dayNames[D + 7],
                    m: m + 1,
                    mm: pad(m + 1),
                    mmm: dF.i18n.monthNames[m],
                    mmmm: dF.i18n.monthNames[m + 12],
                    yy: String(y).slice(2),
                    yyyy: y,
                    h: H % 12 || 12,
                    hh: pad(H % 12 || 12),
                    H: H,
                    HH: pad(H),
                    M: M,
                    MM: pad(M),
                    s: s,
                    ss: pad(s),
                    l: pad(L, 3),
                    L: pad(L > 99 ? Math.round(L / 10) : L),
                    t: H < 12 ? "a" : "p",
                    tt: H < 12 ? "am" : "pm",
                    T: H < 12 ? "A" : "P",
                    TT: H < 12 ? "AM" : "PM",
                    Z: utc ? "UTC" : (String(date).match(timezone) || [""]).pop().replace(timezoneClip, ""),
                    o: (o > 0 ? "-" : "+") + pad(Math.floor(Math.abs(o) / 60) * 100 + Math.abs(o) % 60, 4),
                    S: ["th", "st", "nd", "rd"][d % 10 > 3 ? 0 : (d % 100 - d % 10 != 10) * d % 10]
                };

            return mask.replace(token, function ($0) {
                return $0 in flags ? flags[$0] : $0.slice(1, $0.length - 1);
            });
        };
    }();

    // Some common format strings
    dateFormat.masks = {
        "default": "ddd mmm dd yyyy HH:MM:ss",
        shortDate: "m/d/yy",
        mediumDate: "mmm d, yyyy",
        longDate: "mmmm d, yyyy",
        fullDate: "dddd, mmmm d, yyyy",
        shortTime: "h:MM TT",
        mediumTime: "h:MM:ss TT",
        longTime: "h:MM:ss TT Z",
        isoDate: "yyyy-mm-dd",
        isoTime: "HH:MM:ss",
        isoDateTime: "yyyy-mm-dd'T'HH:MM:ss",
        isoUtcDateTime: "UTC:yyyy-mm-dd'T'HH:MM:ss'Z'"
    };

    // Internationalization strings
    dateFormat.i18n = {
        dayNames: [
            "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
        ],
        monthNames: [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            "January", "February", "March", "April", "May", "June", "July", "August", "September",
            "October", "November", "December"
        ]
    };

    // For convenience...
    Date.prototype.format = function (mask, utc) {
        return dateFormat(this, mask, utc);
    };



    var rooms = [{
            "room": "chime",
            "name": "cr.pneb2.2.chime.4",
            "address": "cr.pneb2.2.chime.4@symantec.com"
        },
        {
            "room": "clarinet",
            "name": "CR.PNEB2.2.Clarinet.10",
            "address": "cr.pneb2.2.clarinet.10@symantec.com"
        },
        {
            "room": "conga",
            "name": "CR.PNEB2.2.Conga.8",
            "address": "cr.pneb2.2.conga.8@symantec.com"
        },
        {
            "room": "melodica",
            "name": "CR.PNEB2.2.Melodica.16(VC)",
            "address": "cr.pneb2.2.melodica.16@symantec.com"
        }
    ];

    let parms = {
        title: 'event',
        active: {
            event: true
        }
    };

    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;

    if (accessToken && userName) {
        parms.user = userName;

        // Initialize Graph client
        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        try {


            // Get the first 10 contacts in alphabetical order
            // by given name
            console.log('----- String ---' + req.param('person'));
            var reserveKey = ["with", "subject", "body", "room", "time"];
            var keyPoint = new Object();
            var finalMap = new Object();
            var myString = req.param('person');
            var stage = req.param('stage');

            if (stage == 'Initial') {
                this.mailto = null;
                this.mailsubject = null;
                this.mailbody = null;
                this.personName = null;
                this.roomadd = null;
                this.roomname = null;
                this.starttime = null;
                this.endtime = null;
            }

            console.log('----stage---' + stage);
            for (var i = 0; i < reserveKey.length; i++) {
                if (myString.lastIndexOf(reserveKey[i]) != -1) {
                    keyPoint[reserveKey[i]] = myString.lastIndexOf(reserveKey[i]) + reserveKey[i].length;
                }
            }


            var map1 = sortProperties(keyPoint);
            for (var key in map1) {
                finalMap[String(map1[key]).split(',')[0]] = myString.slice(String(map1[key]).split(',')[1], (map1[++key] == undefined ? myString.length : String(map1[key]).split(',')[1] - String(map1[key]).split(',')[0].length));
            }

            var resultData = '<Html><table style="width:100%;border:1px solid black;">';
            var emailSearch = finalMap['with'];

            console.log('-- email search--' + emailSearch);


            //var mailto;
            //var mailsubject;
            //var mailbody;
            var bobmsg;

            if (emailSearch !== undefined && emailSearch != ' ') {

                console.log('-- email search 1--' + emailSearch);
                const result = await client
                    .api('/me/people/?$search=' + emailSearch)
                    .version("beta")
                    .top(1)
                    .get();

                if (result.value[0] !== undefined) {
                    console.log('--out mail --' + stage);
                    console.log('--out mail 1--' + this.mailto);
                    if (result.value.length > 1) {
                        for (i = 0; i < result.value.length; i++) {
                            console.log('-- search result-->' + result.value[i].displayName);
                            if ((result.value[i].displayName).includes(emailSearch)) {
                                console.log('-- search result-->' + result.value[i].userPrincipalName);
                                this.personName = result.value[i].displayName;
                                this.mailto = result.value[i].userPrincipalName;
                            }

                        }
                    } else {
                        this.personName = result.value[0].displayName;
                        this.mailto = result.value[0].userPrincipalName;
                    }


                    console.log('--in mail 1--');
                    //this.mailto = result.value[0].userPrincipalName;
                    console.log('--in mail 1--' + this.mailto);
                    //this.personName = result.value[0].displayName;	    

                    //resultData+= '<tr><td>To:</td><td>'+result.value[0].userPrincipalName+'</td></tr>';
                }
                //console.log('---->'+result.value[0].userPrincipalName);	    
            }






            /* const result1 = await client
             .api('/me/events')
             .post(event, (err, res) => {
               console.log(JSON.stringify(err)+'Event Response -> '+JSON.stringify(res));
              });*/

            console.log('wow-- this.mailto--' + this.mailto);

            if (this.mailto != undefined) {
                resultData += '<tr><td>To:</td><td>' + this.mailto + '</td></tr>';
            } else {
                bobmsg = 'Tell me email address';
            }

            console.log('wow-- this.mailsubject--' + this.mailsubject);

            if (finalMap['subject'] != undefined) {
                this.mailsubject = finalMap['subject'];
            }
            console.log('wow-- this.room--' + finalMap['room']);
            if (finalMap['room'] != undefined) {
                for (i = 0; i < rooms.length; i++) {
                    console.log(finalMap['room'] + '----wow-- rooms[i].name--' + rooms[i].room);
                    console.log(finalMap['room'].length + '----wow-- rooms[i].name--' + (rooms[i].room).length);

                    if (' ' + rooms[i].room === finalMap['room']) {
                        console.log('-- in room--' + rooms[i].address);
                        this.roomadd = rooms[i].address;
                        this.roomname = rooms[i].name;
                    }
                }
            }



            if (this.mailsubject != undefined) {
                console.log('--out subject 1--' + stage);
                console.log('--out subject 2--' + finalMap['subject']);
                resultData += '<tr><td>Subject:</td><td>' + this.mailsubject + '</td></tr>';
                console.log('--in subject 2--' + resultData);
            } else {
                bobmsg = 'Please Help me with Subject line.'
            }

            if (finalMap['body'] != undefined) {
                this.mailbody = finalMap['body'];
            }
            if (this.mailbody != undefined) {
                resultData += '<tr><td>Body:</td><td>' + this.mailbody + '</td></tr>';
            }

            if (this.roomadd != undefined) {
                resultData += '<tr><td>Conference Room:</td><td>' + this.roomname + '</td></tr>';
                console.log('--room--' + this.roomname);
            } else {
                if (bobmsg == undefined) {
                    bobmsg = 'Please Select Conference Room for meeting.';
                }
            }

            if (finalMap['time'] != undefined) {
                console.log('----- time ----' + finalMap['time']);
                var splitData = finalMap['time'].substr(1).split(' ');
                var day = splitData[0].match(/\d+/g).map(Number);
                var month = splitData[1];
                var year = (new Date()).getFullYear();
                var fromTime, toTime;
                if (splitData[2].toUpperCase().indexOf('P') != -1)
                    fromTime = 12 + Number(splitData[2].match(/\d+/g).map(Number));
                else
                    fromTime = splitData[2].match(/\d+/g).map(Number);

                if (splitData[4].toUpperCase().indexOf('P') != -1)
                    toTime = 12 + Number(splitData[4].match(/\d+/g).map(Number));
                else
                    toTime = splitData[5].match(/\d+/g).map(Number);
                var now = new Date(month + ' ' + day + ', ' + year + ' ' + fromTime + ':00:00');
                this.starttime = now.format("isoDateTime");
                now = new Date(month + ' ' + day + ', ' + year + ' ' + toTime + ':00:00');
                this.endtime = now.format("isoDateTime");
                console.log(this.starttime + '----time---' + this.endtime);
            }
            if (this.starttime != undefined) {
                resultData += '<tr><td>Start Time:</td><td>' + this.starttime + '</td></tr>';
                resultData += '<tr><td>End Time:</td><td>' + this.endtime + '</td></tr>';
            } else {
                if (this.roomadd != undefined) {
                    bobmsg = 'Please specify Start time and End time.';
                }
            }

            var event = {
                "subject": (this.mailsubject != undefined ? this.mailsubject : ''),
                "body": {
                    "contentType": "HTML",
                    "content": (this.mailbody != undefined ? this.mailbody : '')
                },
                "start": {
                    "dateTime": (this.starttime != undefined ? this.starttime : ''),
                    "timeZone": "Pacific Standard Time"
                },
                "end": {
                    "dateTime": (this.endtime != undefined ? this.endtime : ''),
                    "timeZone": "Pacific Standard Time"
                },
                "location": {
                    "displayName": (this.roomname != undefined ? this.roomname : '')
                },
                "attendees": [{
                        "emailAddress": {
                            "address": (this.mailto != undefined ? this.mailto : ''),
                            "name": (this.personName != undefined ? this.personName : '')
                        },
                        "type": "required"

                    },
                    {
                        "emailAddress": {
                            "address": (this.roomadd != undefined ? this.roomadd : ''),
                            "name": (this.roomname != undefined ? this.roomname : '')
                        }
                    }
                ]
            }

            console.log('Event Json----->' + JSON.stringify(event));
            console.log(' captain america ----->');

            // Meeting Booking Validation
            var postDataJSON = '{ "attendees": [ { "type": "required", "emailAddress": { "address": "' + this.mailto + '" } } ], "locationConstraint": { "isRequired": "false", "suggestLocation": "false", "locations": [ { "resolveAvailability": "false", "locationEmailAddress": "' + this.roomadd + '" } ] }, "timeConstraint": { "activityDomain":"work", "timeslots": [ { "start": { "dateTime": "' + this.starttime + '", "timeZone": "UTC" }, "end": { "dateTime": "' + this.endtime + '", "timeZone": "UTC" } } ] }, "meetingDuration": "PT60M", "returnSuggestionReasons": "false", "minimumAttendeePercentage": "100" }';
            console.log(' Iron Man ----->');

            const meetingResult = await client
                .api('/me/findMeetingTimes')
                .version("beta")
                .post(postDataJSON, (err, meetingResult) => {
                    
                    console.log('----- Hulk -----');
              console.log('---- Result meetingResult ----->' + meetingResult);

            if (meetingResult.emptySuggestionsReason !== undefined) {
                if (meetingResult.emptySuggestionsReason == '') { // Positive Response Available From Server
                    console.log(meetingResult);
                    if (stage == 'ready to send' && (myString === 'send' || myString === 'yes')) {
                        const result1 = await client
                            .api('/me/events')
                            .post(event, (err, res) => {
                                console.log(JSON.stringify(err) + 'Event Response -> ' + JSON.stringify(res));
                            });

                        bobmsg = 'meeting set successfully with ' + this.personName + '. Have a good day';
                        stage = 'Initial';
                        this.mailto = null;
                        this.mailbody = null;
                        this.mailsubject = null;
                        this.starttime = null;
                        this.endtime = null;
                        this.roomname = null;
                        this.roomadd = null;

                        resultData = 'Meeting Set Successfully';
                    }
                } else {
                    if (bobmsg == undefined) {
                        bobmsg = 'Attendees unavailable at this time';
                    }
                }
            }else{
                console.log('-----availibility else-----');
            }
             
              });
            
            
              

            if (bobmsg == undefined) {
                console.log('---- bob ---ready');
                bobmsg = 'Mail is ready to Send. Are you sure you want to send ?';
                stage = 'ready to send';
            }


            resultData += '</table></html>';
            if (stage == 'Initial') {
                stage = 'in progress';
            }
            res.status(200).json({
                bob: bobmsg,
                consoleoutput: resultData,
                state: stage
            });

        } catch (err) {
            console.log('--err---' + err.message);
            console.log('--err stack--' + err.stack);
            parms.message = 'Error retrieving contacts';
            parms.error = {
                status: `${err.code}: ${err.message}`
            };
            parms.debug = JSON.stringify(err.body, null, 2);
            res.render('error', parms);
        }

    } else {
        // Redirect to home
        res.redirect('/');
    }
});

module.exports = router;
