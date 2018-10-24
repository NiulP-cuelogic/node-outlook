var express = require('express');
var app = express();

var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');
var session = require('express-session');
var moment = require('moment');
// var querystring = require('querystring');
var outlook = require('node-outlook');
var pages = require('./pages');
var authHelper = require('./authHelper');

app.use(express.static('static'));

app.use(bodyParser.json());

app.use(cookieParser());
app.use(session(
  { secret: '0dc529ba-5051-4cd6-8b67-c9a901bb8bdf',
    resave: false,
    saveUninitialized: false 
  }));
  

app.get('/', function(req, res) {
  res.send(pages.loginPage(authHelper.getAuthUrl()));
});

app.get('/authorize', function(req, res) {
  var authCode = req.query.code;
  if (authCode) {
    console.log('');
    console.log('Retrieved auth code in /authorize: ' + authCode);
    authHelper.getTokenFromCode(authCode, tokenReceived, req, res);
  }
  else {
    console.log('/authorize called without a code parameter, redirecting to login');
    res.redirect('/');
  }
});

function tokenReceived(req, res, error, token) {
  if (error) {
    console.log('ERROR getting token:'  + error);
    res.send('ERROR getting token: ' + error);
  }
  else {
    req.session.access_token = token.token.access_token;
    req.session.refresh_token = token.token.refresh_token;
    req.session.email = authHelper.getEmailFromIdToken(token.token.id_token);
    res.redirect('/logincomplete');
  }
}

app.get('/logincomplete', function(req, res) {
  var access_token = req.session.access_token;
  var refresh_token = req.session.access_token;
  var email = req.session.email;
  
  if (access_token === undefined || refresh_token === undefined) {
    console.log('/logincomplete called while not logged in');
    res.redirect('/');
    return;
  }
  
  res.send(pages.loginCompletePage(email));
});

app.get('/refreshtokens', function(req, res) {
  var refresh_token = req.session.refresh_token;
  if (refresh_token === undefined) {
    console.log('no refresh token in session');
    res.redirect('/');
  }
  else {
    authHelper.getTokenFromRefreshToken(refresh_token, tokenReceived, req, res);
  }
});

app.get('/logout', function(req, res) {
  req.session.destroy();
  res.redirect('/');
});

app.get('/sync', function(req, res) {
  var token = req.session.access_token;
  var email = req.session.email;
  if (token === undefined || email === undefined) {
    console.log('/sync called while not logged in');
    res.redirect('/');
    return;
  }
  
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
  
  outlook.base.setAnchorMailbox(req.session.email);

  outlook.base.setPreferredTimeZone('Eastern Standard Time');
  
  var requestUrl = req.session.syncUrl;
  if (requestUrl === undefined) {
    requestUrl = outlook.base.apiEndpoint() + '/Me/CalendarView';
  }
  
  var startDate = moment().startOf('day');
  var endDate = moment(startDate).add(7, 'days');
  var params = {
    startDateTime: startDate.toISOString(),
    endDateTime: endDate.toISOString()
  };
  
  
  var headers = {
    Prefer: [ 
      'odata.track-changes',
      'odata.maxpagesize=5'
    ]
  };
  
  var apiOptions = {
    url: requestUrl,
    token: token,
    headers: headers,
    query: params
  };
  
  outlook.base.makeApiCall(apiOptions, function(error, response) {
    if (error) {
      console.log(JSON.stringify(error));
      res.send(JSON.stringify(error));
    }
    else {
      if (response.statusCode !== 200) {
        console.log('API Call returned ' + response.statusCode);
        res.send('API Call returned ' + response.statusCode);
      }
      else {
        var nextLink = response.body['@odata.nextLink'];
        if (nextLink !== undefined) {
          req.session.syncUrl = nextLink;
        }
        var deltaLink = response.body['@odata.deltaLink'];
        if (deltaLink !== undefined) {
          req.session.syncUrl = deltaLink;
        }
        res.send(pages.syncPage(email, response.body.value));
      }
    }
  });
});

app.post('/createEvent', function(req, res) {

    // var token = req.session.access_token;
    var token = 'EwAgA+l3BAAUv0lYxoez7x2t6RowHa2liVeLW/wAAVmtzS7D+j4g6vxf3tBqxSGNo5KNO7wTA763j7ejRCcYicOBPkT0j74ozdxEahC5xeoLnm3tPqt56gfMAoOQCC1Q8Ud3Etb7QbmKz1aFbnfctz9JGuv0EKDNTdk6NFQOXbhYP1OjNCum9Eo17Og7P1cDm7ZqgCaH5gQv6TTfKVmHBe4AwSs720+Qwa/bieUOoIwE4Vzs55bTM7rprcEIa9CjOj1+P/UDaoVaHrALu7QfX5WRRiXkZ0d5ISiqQyBIP3TYQ0rvdiNonkPGIsxzZcegefIGIGi8JobpNCO5c01s4YlZbiyCBB48uSW/BwItitSmCs7h9IIRPcOPCDfFVxkDZgAACPG3R5CX9g++8AGBxAt5Vm0r0NRf+P8HZ9P1LNa9fF6+XUCe8tQT+y9KjJ5MqpJYIbHy09uu/yga5UVaz5/WR02bsfug3g9xQeYhvQ4lllZSYvpQaAX3CfBWupTpWrDl6w4g1u85RsH7/fAvqPW3bWbLxGDZ3rPimXBhxg66ntQ7oFDcmHubdARmtFglRmRu58fADHwZ6IqqvgU/vY8qI9Ec/2rI7mTVPEMnF/FclgIyCmUmAa6vT6bYZ2MoX89R1DGG9tewgzKlHrBnPmKYdtxtEhAnRMNsUtKhpkm3DXVmX20KaT3JCMhdYBTo6iVUKJWxOd1G1Z5j5qa5VrAYZgjxyz10WzHtIE3dUeIWtt2Rrc4nNZixS9SyG1zmq8M1V0kkF9Z1XJ4kWEAd/FWJr3kiauA4GzX8fa8izkNlP2sVoMLbwbNLSVdVXvmsgpMtr3LyAe1PeBCsksoGCdeH7rDf8J0Yq3u+S0zzaxwJ1pI8xldM60KpdDxCKr34XL069nBJzEe+Ia/id0eVempVDzIdjf6+ythOtGJ4cAt+nwmclJXujgozkmALIBKRsEvgoW1jfsiq1Szzc8W4CeslDoN4s1LNhTQI/+ogDe3ezq9Ay5aPxX1gOVA0ZLFJ4IZQ2NFlGGEcH6CrsHncTTpSAMRXla1DBRkat0EvIwI='
    // var email = req.session.email;
    var email = 'niul.omega@outlook.com';

    console.log("token=========================================", token);
    console.log("email====================================", email);

    if (token === undefined || email === undefined) {
      console.log('/sync called while not logged in');
      res.redirect('/');
      return;
    }

  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
  
  outlook.base.setAnchorMailbox(email);

  outlook.base.setPreferredTimeZone('Eastern Standard Time');

  var newEvent = {
    Subject: "Discuss the Calendar REST API with deepali ",
    Body : {
    ContentType: "HTML",
    Content: "I think it will meet our requirements!"
  },
    Start: {
      DateTime: "2018-10-29T18:00:00",
      TimeZone: "India Standard Time"
  },
    End: {
      DateTime: "2018-10-29T19:00:00",
      TimeZone: "India Standard Time"
  }
//   Attendees: [
//     {
//       EmailAddress: {
//         Address: "ajit255@hotmail.com",
//         Name: "Janet Schorr"
//       },
//       Type: "Required"
//     }
//   ]
};
    

    var createEventParams = {
        token : token,
        event : newEvent
    };

    outlook.calendar.createEvent(createEventParams, function(error, event) {
        if (error){
            console.log(error);
        }
        else {
            console.log("Event===========================+>",event);
            if (event) {
                requestUrl = outlook.base.apiEndpoint() + '/me/events/'+ event.Id;
                // var apiOptions = {
                //     url: requestUrl,
                //     token: token,
                //     headers: headers,
                //     query: params
                // };
                console.log(event.WebLink);
                res.redirect(event.WebLink);
                // res.end();
            }
        }
    })
})


var server = app.listen(3000, function() {
  var host = server.address().address;
  var port = server.address().port;
  
  console.log('Example app listening at http://%s:%s', host, port);
});