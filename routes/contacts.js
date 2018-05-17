// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* GET /contacts */
router.get('/', async function(req, res, next) {
  let parms = { title: 'Contacts', active: { contacts: true } };

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
      var event = {
    "subject": "Let's go for lunch",
    "body": {
        "contentType": "HTML",
        "content": "Does late morning work for you?"
    },
    "start": {
        "dateTime": "2017-04-15T12:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "end": {
        "dateTime": "2017-04-15T14:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "location": {
        "displayName": "Harry's Bar"
    },
    "attendees": [{
        "emailAddress": {
            "address": "samanthab@contoso.onmicrosoft.com",
            "name": "Samantha Booth"
        },
        "type": "required"
    }]
}
      const result1 = await client
      .api('/me/events')
      .post(event, (err, res) => {
        console.log('Event Response -> '+res);
       })
      console.log('Event Response 1-> '+result1.value);
      const result = await client
      .api('/me/people/?$search=siddh')
      .version("beta")
      .top(1)
      .get();

      parms.contacts = result.value;
      console.log('People--->'+JSON.stringify(result.value));
      console.log('-----------------------------------------------------');
      console.log('People--->'+result.value.emailAddresses);
      res.render('contacts', parms);
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
