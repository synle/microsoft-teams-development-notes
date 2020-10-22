# microsoft-teams-development-notes

### JS API
https://docs.microsoft.com/en-us/javascript/api/overview/msteams-client?view=msteams-client-js-latest

### Tunneling / Reverse port forwarding for local dev
https://medium.com/better-programming/how-to-expose-a-secure-https-url-to-your-local-web-server-eddf780be515

### Sample init snippet
```
npm i --save @microsoft/teams-js
```

#### Initialization
```
function _initMsTeams(): Promise<boolean> {
  return new Promise((resolve) => {
    let isMsTeams = false;

    microsoftTeams.initialize(() => {
      console.log('microsoftTeams.initialize > Done init');
      isMsTeams = true;
      window.microsoftTeams = microsoftTeams;
      resolve(true);
    });

    setTimeout(() => {
      if (isMsTeams === false) {
        resolve(false);
      }
    }, 100);
  }).then((isMsTeams: boolean) => {
    localStorage.setItem('IsFromMSTeam', isMsTeams ? 'true' : 'false');
    return isMsTeams;
  });
}
```

```
_initMsTeams().then(login); // login is defined below
```

#### Trigger auth request from client
```
function login(isFromTeamsApp){
  if(isFromTeamsApp){
    // if it is from teams app, then call the api...
    microsoftTeams.authentication.authenticate({
      url: '/login',
      successCallback: function (...data) {
        console.log('microsoftTeams.initialize.successCallback', data);
        window.location.reload();
      },
      failureCallback: function (...data) {
        console.log('microsoftTeams.initialize.failureCallback', data);
      },
      width: 800,
      height: 600,
    });
  } else {
    // other login flow
  }
}
```


#### Notify auth success
```
// isThisFromTeamsApp is defined by the above call...

async function getUserProfile(){
  try{
    const user = await UserApi.getUserProfile();
    if (isThisFromTeamsApp && window.opener) {
      // this is coming from an authentication popup, then trigger teams authentication and exit
      microsoftTeams.authentication.notifySuccess(user);
      return window.close();
    }
  } catch(err){
    login();// fetch profile failed, then trigger the login flow
  }
}
```
