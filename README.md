# microsoft-teams-development-notes

### Teams App Studio
Download `App Studio` for Teams to assist with manifest editor and tooling

### Create brand new tenant for testing
Good for 90 days

https://cdx.transform.microsoft.com/my-tenants/create-tenant

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
        console.log('microsoftTeams.authentication.successCallback', data);
        window.location.reload();
      },
      failureCallback: function (...data) {
        console.log('microsoftTeams.authentication.failureCallback', data);
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


### Teams SSO with AAD
https://dev.to/urmade/seamless-sso-login-for-microsoft-teams-tabs-3n8k

https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso
