// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* GET /contacts */
router.get('/', async function(req, res, next) {
  let parms = { title: 'Contacts', active: { contacts: true } };
  console.log('---req data--'+req.params.q);
  console.log('---req data 1--'+req.param('q'));
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
      .api('/me/people/?$search=vishal')
      .version("beta")
      .top(1)
      .get();

      parms.contacts = result.value;
      console.log('People--->'+JSON.stringify(result.value));
      console.log('-----------------------------------------------------');
      console.log('People New--->'+result.value[0].userPrincipalName);
      console.log('People New 1--->'+result.value[0].id);
      console.log('People New 2--->'+result.value.id[0]);

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
