const express = require("express");
const msal = require("@azure/msal-node");

const SERVER_PORT = process.env.PORT || 3000;

// Create Express App and Routes
const app = express();

app.listen(SERVER_PORT, () =>
    console.log(
        `Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`
    )
);

const config = {
    auth: {
        clientId: "9ec95365-1690-4c6f-ae51-2cfe1a73ee9e",
        authority: "https://login.microsoftonline.com/common",
        clientSecret: "pOr7Q~aeAWKNoulW0wTrdgI1ZRZKRRUMAmCrI",
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        },
    },
};

// Create msal application object
const cca = new msal.ConfidentialClientApplication(config);

app.get("/", (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read"],
        redirectUri: "http://localhost:3000/auth",
    };

    // get url to sign user in and consent to scopes needed for application
    cca.getAuthCodeUrl(authCodeUrlParameters)
        .then((response) => {
            res.redirect(response);
        })
        .catch((error) => console.log(JSON.stringify(error)));
});

app.get("/auth", (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: "http://localhost:3000/auth",
    };

    cca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            console.log("\nResponse: \n:", response);
            res.sendStatus(200);
        })
        .catch((error) => {
            console.log(error);
            res.status(500).send(error);
        });
});