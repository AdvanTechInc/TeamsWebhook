const dotenv = require('dotenv').config()
const express = require('express')
const http = require('http')
const https = require('https')
const app = express()

const { IncomingWebhook } = require('ms-teams-webhook')

// Initialize IncomingWebhook object with the url of the Teams Webhook information
const url = process.env.MS_TEAMS_WEBHOOK_URL;
const webhook = new IncomingWebhook(url);

// Set view engine and app options
app.set("view engine", "ejs")
app.use(express.json())

// Router entry for webpage
app.get("/", (req, res) => {
    res.render('index', {
        title: 'Teams Webhook'
    })
})

// Router entry for Webhook function
app.post('/json', async (req, res) => {
    // Log the information sent to the webhook from the Webpage
    console.log(req.body.productNum)
    console.log(req.body.ticketNum)
    console.log(req.body.quantity)
    const options = {
        "method": "GET",
        "hostname": "api-na.myconnectwise.net",
        "port": null,
        "path": `/v4_6_release/apis/3.0/procurement/catalog?conditions=identifier=\"${req.body.productNum}\"`,
        "headers": {
          "Accept": "*/*",
          "clientID": `${process.env.CW_CLIENT}`,
          "Content-Type": "application/json",
          "Authorization": `Basic ${process.env.CW_AUTH}`
        }
    }
    const cwReq = https.request(options, function (cwRes) {
        const chunks = [];
      
        cwRes.on("data", function (chunk) {
          chunks.push(chunk);
        });
      
        cwRes.on("end", function () {
          const body = Buffer.concat(chunks);
          console.log(body.toString());
        });
      });
      
      
    // Sends information entered into the site to the webhook
    await webhook.send(JSON.stringify(
        {
            "type": "message",
            "attachments": [{
                "contentType":"application/vnd.microsoft.card.adaptive",
                "contentUrl":null,
                "type": "AdaptiveCard",
                "content":{
                "body": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "Product Checkout"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "Image",
                                        "style": "Person",
                                        "url": "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
                                        "size": "Small"
                                    }
                                ],
                                "width": "auto"
                            },
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "weight": "Bolder",
                                        "text": `${req.body.submittedBy}`,
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "spacing": "None",
                                        "text": "Created",
                                        "isSubtle": true,
                                        "wrap": true
                                    }
                                ],
                                "width": "stretch"
                            }
                        ]
                    },
                    {
                        "type": "FactSet",
                        "facts": [
                            {
                                "title": "Ticket:",
                                "value": `[${req.body.ticketNum}](https://na.myconnectwise.net/v4_6_release/services/system_io/Service/fv_sr100_request.rails?service_recid=${req.body.ticketNum}&companyName=advantechinc)`
                            },
                            {
                                "title": "Product Name:",
                                "value": `[${req.body.productNum}]()`
                            },
                            /*{
                                "title": "Product Description:",
                                "value": `${cwRes.body.description}`
                            },*/
                            {
                                "title": "Product Number:",
                                "value": `${req.body.productNum}`
                            },
                            {
                                "title": "Quantity:",
                                "value": `${req.body.quantity}`
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            }
            }]
        }
    )).then(() => console.log("Message Sent")).catch(err => console.log(err))
    cwReq.end();
})



// Creates web server for the express app
http.createServer(app).listen(8080, () => {
    console.log('HTTP server up at port 8080.')
    console.log(url)
})