// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

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
    console.log('----- event email ---'+req.param('person'));
    var emailSearch = req.param('person').split("with");
    const result = await client
      .api('/me/people/?$search='+emailSearch[1])
      .version("beta")
      .top(1)
      .get(); 
    console.log('----searched Email------>'+result.value[0].userPrincipalName);  
      
    var event = {
    "subject": "Test mail by BOT",
    "body": {
        "contentType": "HTML",
        "content": "Does late morning work for you?"
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
      
      
      const result1 = await client
      .api('/me/events')
      .post(event, (err, res) => {
        console.log(JSON.stringify(err)+'Event Response -> '+JSON.stringify(res));
       });
      
      
      res.redirect('/');
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
