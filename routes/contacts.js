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
  
      const result = await client
      .api('/me/people/?$search=priyanka')
      .version("beta")
      .top(100)
      .get();

      parms.contacts = result.value;
      console.log('People--->'+JSON.stringify(result.value));
      console.log('-----------------------------------------------------');
      
      //res.render('index', parms);
      var resultData = '<Html><table style="width:100%;border:1px solid black;"><tr><td>To:</td><td>'+result.value[0].userPrincipalName+'</td></tr></table></html>';
      res.status(200).json(JSON.stringify(result.value));

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
