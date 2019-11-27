import {inject, Aurelia} from 'aurelia-framework';
import {HttpClient} from 'aurelia-fetch-client';

@inject(Aurelia, HttpClient)
export class Login {
  message = 'Login Page';

  msalConfig = {
    auth: {
      clientId: '<client id>',
      authority: 'https://login.microsoftonline.com/<tenancy>'
    },
    cache: {
      storeAuthStateInCookie: true,
      cacheLocation: 'localStorage'
    }
  };
  loginRequest = {
    scopes: ['user.read.all']
  };
  graphEndpoint = 'https://graph.microsoft.com/v1.0/me';

  myMSALObj = new Msal.UserAgentApplication(this.msalConfig);
  graphData;

  constructor(aurelia, httpClient) {
    this.aurelia = aurelia;
    httpClient.configure(config => {
      config
        .useStandardConfiguration();
    });
    this.httpClient = httpClient;
  }

  attached() {
    this.myMSALObj.handleRedirectCallback(this.authCallback);
    if (this.myMSALObj.getAccount()) {
      this.acquireTokenPopupAndCallMSGraph();
    } else {
      this.myMSALObj.loginRedirect(this.loginRequest);
    }
  }

  authCallback(error, response) {
    localStorage.setItem('login_response', JSON.stringify(response));
    this.acquireTokenPopupAndCallMSGraph();
  }

  acquireTokenPopupAndCallMSGraph() {
    this.myMSALObj.acquireTokenSilent(this.loginRequest)
      .then(accessTokenResponse => {
        localStorage.setItem('access_token_response', JSON.stringify(accessTokenResponse));
        this.callMSGraph(this.graphEndpoint, accessTokenResponse.accessToken);
      })
      .catch(error => {
        if (error.name === 'InteractionRequiredAuthError') {
          this.myMSALObj.acquireTokenRedirect(this.loginRequest)
            .then(accessToken => {
              localStorage.setItem('access_token_response', JSON.stringify(accessTokenResponse));
              this.callMSGraph(this.graphEndpoint, accessToken);
            });
        }
      });
  }

  callMSGraph(theUrl, accessToken) {
    localStorage.setItem('access_token', accessToken);
    this.httpClient.fetch(theUrl, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      }
    })
      .then(response => response.json())
      .then(fetchedData => {
        this.graphAPICallback(fetchedData);
      });
    this.httpClient.fetch("https://graph.microsoft.com/v1.0/users?$filter=startswith(jobTitle,'User') or startswith(jobTitle,'Administrator')&$select=displayName", {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      }
    })
      .then(response => response.json())
      .then(fetchedUsers => {
        localStorage.setItem('users', JSON.stringify(fetchedUsers.value));
      });
  }

  graphAPICallback(data) {
    localStorage.setItem('profile', JSON.stringify(data));
    this.aurelia.setRoot(PLATFORM.moduleName('app'));
  }
}
