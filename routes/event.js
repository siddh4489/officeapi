// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

function sortProperties(obj)
{
  // convert object into array
	var sortable=[];

	for(var key in obj)
		if(obj.hasOwnProperty(key))
			sortable.push([key, obj[key]]); // each item is an array in format [key, value]
	
	// sort items by value
	sortable.sort(function(a, b)
	{
		var x=a[1],
			y=b[1];
		return x<y ? -1 : x>y ? 1 : 0;
	});
	return sortable; // array in format [ [ key1, val1 ], [ key2, val2 ], ... ]
}

/* GET /contacts */
router.get('/', async function(req, res, next) {
  
  
  
  let parms = { title: 'event', active: { event: true } };

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
    console.log('----- String ---'+req.param('person'));
    var reserveKey = ["with","subject","body","room","time"];
    var keyPoint = new Object();
    var finalMap = new Object();
    var myString = req.param('person');
for(var i=0;i<reserveKey.length;i++){
    if(myString.lastIndexOf(reserveKey[i]) != -1){
       keyPoint[reserveKey[i]] = myString.lastIndexOf(reserveKey[i])+reserveKey[i].length;
    }
}


var map1 = sortProperties(keyPoint);    
   for(var key in map1){
     finalMap[String(map1[key]).split(',')[0]] = myString.slice(String(map1[key]).split(',')[1],(map1[++key]==undefined?myString.length:String(map1[key]).split(',')[1]-String(map1[key]).split(',')[0].length));
    }  

  console.log(' Name '+finalMap['with']);
  console.log(' Subject '+finalMap['subject']);	
	console.log(' Body '+finalMap['body']);	     
	    console.log(' room '+finalMap['room']);	
	      console.log(' time '+finalMap['time']);	    




	    
	    
	    
    var emailSearch = finalMap['with'];
    const result = await client
      .api('/me/people/?$search='+emailSearch[1])
      .version("beta")
      .top(1)
      .get(); 
    console.log('----searched Email------>'+result.value[0].userPrincipalName);  
      
    var event = {
    "subject": finalMap['subject'],
    "body": {
        "contentType": "HTML",
        "content": finalMap['body']
    },
    "start": {
        "dateTime": "2018-06-04T12:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "end": {
        "dateTime": "2018-06-04T14:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "location": {
        "displayName": "CR.PNEB2.2.Chime.4"
    },
    "attendees": [{
        "emailAddress": {
            "address": result.value[0].userPrincipalName,
            "name": result.value[0].displayName
        },
        "type": "required"
    
    }]
}
    
    console.log('Event Json----->'+JSON.stringify(event));
      
      
     /* const result1 = await client
      .api('/me/events')
      .post(event, (err, res) => {
        console.log(JSON.stringify(err)+'Event Response -> '+JSON.stringify(res));
       });*/
      
      var resultData = '<Html><table style="width:100%;border:1px solid black;">';
	  resultData+= '<tr><td>To:</td><td>'+result.value[0].userPrincipalName+'</td></tr>';
	  resultData+= '<tr><td>Subject:</td><td>'+finalMap['body']+'</td></tr>';
	  resultData+= '<tr><td>Body:</td><td>'+finalMap['subject']+'</td></tr>'; 
	  resultData+= '</table></html>';  
	    
      res.status(200).json(resultData);	    
      
     // res.redirect('/');
    } catch (err) {
      parms.message = 'Error retrieving contacts';
      parms.error = { status: `${err.code}: ${err.message}` };
      parms.debug = JSON.stringify(err.body, null, 2);
      res.render('error', parms);
    }
    
  } else {
    // Redirect to home
    res.redirect('/');
  }
});

module.exports = router;
