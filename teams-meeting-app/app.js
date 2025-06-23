// MSAL configuration
const msalConfig = {
    auth: {
        clientId: "120b3e6e-1d37-4fd3-b59e-f99a895ab58f", // IMPORTANT: Replace with your app's Client ID
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin + '/teams-meeting-app/index.html'
    },
    cache: {
        cacheLocation: "sessionStorage", // This is more secure than localStorage
        storeAuthStateInCookie: false,
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Define the scopes needed for Graph API
const loginRequest = {
    scopes: ["User.Read", "Calendars.ReadWrite"]
};

let accountId = null;

function updateUI() {
    const account = msalInstance.getActiveAccount();
    const signInButton = document.getElementById('signInButton');
    const signOutButton = document.getElementById('signOutButton');
    const meetingForm = document.getElementById('meeting-form');
    const userNameSpan = document.getElementById('userName');

    if (account) {
        accountId = account.homeAccountId;
        userNameSpan.innerText = account.username;
        signInButton.style.display = 'none';
        signOutButton.style.display = 'block';
        meetingForm.style.display = 'block';
    } else {
        signInButton.style.display = 'block';
        signOutButton.style.display = 'none';
        meetingForm.style.display = 'none';
    }
}

async function signIn() {
    try {
        await msalInstance.loginRedirect(loginRequest);
    } catch (error) {
        console.error(error);
        updateResponse(`Error during sign-in: ${error.message}`, true);
    }
}

function signOut() {
    const logoutRequest = {
        account: msalInstance.getAccountByHomeId(accountId)
    };
    msalInstance.logoutRedirect(logoutRequest);
}

async function createMeeting(event) {
    event.preventDefault();

    const account = msalInstance.getAllAccounts()[0];
    if (!account) {
        updateResponse("You must be signed in to create a meeting.", true);
        return;
    }

    const tokenRequest = {
        scopes: ["Calendars.ReadWrite"],
        account: account
    };

    try {
        const response = await msalInstance.acquireTokenSilent(tokenRequest);
        const accessToken = response.accessToken;

        const subject = document.getElementById('subject').value;
        const attendees = document.getElementById('attendees').value.split(',').map(email => email.trim()).filter(email => email);
        const startDateTime = new Date(document.getElementById('startDateTime').value).toISOString();
        const endDateTime = new Date(document.getElementById('endDateTime').value).toISOString();

        const meeting = {
            subject: subject,
            start: {
                dateTime: startDateTime,
                timeZone: "UTC"
            },
            end: {
                dateTime: endDateTime,
                timeZone: "UTC"
            },
            attendees: attendees.map(email => ({
                emailAddress: { address: email },
                type: "required"
            })),
            isOnlineMeeting: true,
            onlineMeetingProvider: "teamsForBusiness"
        };

        const graphResponse = await callGraphApi(accessToken, "https://graph.microsoft.com/v1.0/me/events", meeting);

        if (graphResponse.id) {
            updateResponse(`Meeting successfully created! <a href="${graphResponse.webLink}" target="_blank">View meeting</a>`, false, true);
        } else {
            updateResponse("Error creating meeting. Check console for details.", true);
        }

    } catch (error) {
        console.error(error);
        if (error instanceof msal.InteractionRequiredAuthError) {
            try {
                // Fallback to interactive token acquisition
                const response = await msalInstance.acquireTokenPopup(tokenRequest);
                // Retry createMeeting after getting token
            } catch (popupError) {
                console.error(popupError);
                updateResponse(`Error acquiring token: ${popupError.message}`, true);
            }
        } else {
            updateResponse(`Error: ${error.message}`, true);
        }
    }
}

async function callGraphApi(accessToken, endpoint, payload) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

    const options = {
        method: payload ? "POST" : "GET",
        headers: headers,
        body: payload ? JSON.stringify(payload) : null
    };

    const response = await fetch(endpoint, options);
    return await response.json();
}


function updateResponse(message, isError = false, isHtml = false) {
    const responseDiv = document.getElementById('response');

    // Clear previous content
    responseDiv.innerHTML = '';

    if (isHtml) {
        // For the success message which contains a link
        const successContainer = document.createElement('div');
        successContainer.innerHTML = message; // Safely parse the HTML string
        responseDiv.appendChild(successContainer);
    } else {
        // For all other messages, treat as plain text
        responseDiv.textContent = message;
    }

    responseDiv.className = isError ? 'error' : 'success';
}


// Event Listeners
document.getElementById('createMeetingForm').addEventListener('submit', createMeeting);

// Handle the redirect promise. This should be called on every page load.
msalInstance.handleRedirectPromise()
    .then((response) => {
        if (response) {
            // If we have a response, an account was just authenticated. Set it as active.
            msalInstance.setActiveAccount(response.account);
        }
        updateUI();
    })
    .catch((error) => {
        console.error(error);
        updateResponse(`Error processing login response: ${error.message}`, true);
    }); 