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
//var bobmsg;
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
  
  var rooms =[{"room":"chime",
  "name":"cr.pneb2.2.chime.4",
  "address":"cr.pneb2.2.chime.4@symantec.com"
  },
  {"room":"clarinet",
  "name":"CR.PNEB2.2.Clarinet.10",
  "address":"cr.pneb2.2.clarinet.10@symantec.com"
  },
  {"room":"clarinet",
  "name":"CR.PNEB2.2.Conga.8",
  "address":"cr.pneb2.2.conga.8@symantec.com"
  },
  {"room":"Melodica",
  "name":"CR.PNEB2.2.Melodica.16(VC)",
  "address":"cr.pneb2.2.melodica.16@symantec.com"
  }
  ];
  
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
    var stage =req.param('stage');
	  
    console.log('----stage---'+stage);	    
for(var i=0;i<reserveKey.length;i++){
    if(myString.lastIndexOf(reserveKey[i]) != -1){
       keyPoint[reserveKey[i]] = myString.lastIndexOf(reserveKey[i])+reserveKey[i].length;
    }
}


var map1 = sortProperties(keyPoint);    
   for(var key in map1){
     finalMap[String(map1[key]).split(',')[0]] = myString.slice(String(map1[key]).split(',')[1],(map1[++key]==undefined?myString.length:String(map1[key]).split(',')[1]-String(map1[key]).split(',')[0].length));
    }  
	    
    var resultData = '<Html><table style="width:100%;border:1px solid black;">';	    
    var emailSearch = finalMap['with'];
   	    
    console.log('-- email search--'+emailSearch);
    
    
    //var mailto;
    //var mailsubject;
    //var mailbody;
    var bobmsg;	 
       
    if(emailSearch !== undefined && emailSearch != ' '){
	    
      console.log('-- email search 1--'+emailSearch);
       const result = await client	    	    
      .api('/me/people/?$search='+emailSearch)
      .version("beta")
      .top(1)
      .get(); 
	
      if(result.value[0] !== undefined){
	   console.log('--out mail --'+stage);
	   console.log('--out mail 1--'+this.mailto);
	      if(result.value.length>1){
	      	for(i = 0; i< result.value.length; i++){    
			console.log('-- search result-->'+result.value[i].displayName);
			if((result.value[i].displayName).includes(emailSearch)){
			        console.log('-- search result-->'+result.value[i].userPrincipalName);
				this.personName = result.value[i].displayName;
				this.mailto = result.value[i].userPrincipalName;
		        }
		           
		} 
	      }else{
		      this.personName = result.value[0].displayName;
		      this.mailto = result.value[0].userPrincipalName;
		 }
	        

		   console.log('--in mail 1--');
		   //this.mailto = result.value[0].userPrincipalName;
		   console.log('--in mail 1--'+this.mailto);
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
	    
         //console.log('-- email result--'+result);	    
          if(this.mailto != undefined){
           	resultData+= '<tr><td>To:</td><td>'+this.mailto+'</td></tr>';
	   }else{
          	bobmsg ='Tell me email address';
	   }
	  if(finalMap['subject'] != undefined){
	  	this.mailsubject=finalMap['subject']; 
	  }
	  if(finalMap['room'] != undefined){
		  
		 for(i = 0; i< rooms.length; i++){    
		 	if(rooms[i].name.includes(finalMap['room']) || rooms[i].name === (finalMap['room'])){
				this.roomadd=rooms[i].address;
				this.roomname=rooms[i].name;
		       }
		 } 
	  }  
	  
	  
	    
	  if(this.mailsubject != undefined){
	        console.log('--out subject 1--'+stage);  
                console.log('--out subject 2--'+finalMap['subject']);    
	  		resultData+= '<tr><td>Subject:</td><td>'+this.mailsubject+'</td></tr>';
			console.log('--in subject 2--'+resultData);
	  }else{
		bobmsg =  'Please Help me with Subject line. It is required'
	  }
	  if(this.roomadd != undefined){
	  		resultData+= '<tr><td>Conference Room:</td><td>'+this.roomname+'</td></tr>';
			console.log('--room--'+this.roomname);
	  }else{
		bobmsg =  'Please Select Conference Room for meeting.'
	  }    
	  if(finalMap['body'] != undefined){
	  	this.mailbody=finalMap['body']; 
	  }  
	  if(this.mailbody != undefined){
	  	resultData+= '<tr><td>Body:</td><td>'+this.mailbody+'</td></tr>'; 
	  }
	    
	   var event = {
    "subject": (this.mailsubject != undefined?this.mailsubject:''),
    "body": {
        "contentType": "HTML",
        "content": (this.mailbody != undefined?this.mailbody:'')
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
            "address": (this.mailto != undefined?this.mailto:''),
            "name": (this.personName != undefined?this.personName:'')
        },
        "type": "required"
    
    }]
}
    
    console.log('Event Json----->'+JSON.stringify(event));  
	    
	   if(stage == 'ready to send' && (myString === 'send' || myString === 'yes')){
		  const result1 = await client
		      .api('/me/events')
		      .post(event, (err, res) => {
			console.log(JSON.stringify(err)+'Event Response -> '+JSON.stringify(res));
		       }); 
		   
		  bobmsg ='meeting set successfully with '+this.personName+'. Have a good day';
		  stage = 'Initial';
		  this.mailto = null;
		  this.mailbody =null;
		  this.mailsubject =null;
		  resultData = 'Meeting Set Successfully'; 
	     } 
	  if(bobmsg == undefined){
		bobmsg ='Mail is ready to Send. Are you sure you want to send ?';  
		stage = 'ready to send';
	  }
	    
	    
	  resultData+= '</table></html>';  
	      if(stage =='Initial'){
	       stage='in progress'; 
	     }	    
      res.status(200).json({bob:bobmsg,consoleoutput:resultData,state:stage});	    
      
    } catch (err) {
      console.log('--err---'+err.message);  
      console.log('--err stack--'+err.stack);	    
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
