/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

// Bot Framework Dependencies

var restify = require('restify');
var builder = require('botbuilder');

// DocXTemplater Dependencies

var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');

// Nodemailer
var nodemailer = require('nodemailer');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    // appId: process.env.MicrosoftAppId,
    appId: '33d99977-7df3-4f5d-b771-ad32c9ed590a',
    // appPassword: process.env.MicrosoftAppPassword,
    appPassword: "1zf9ijdJWHBqnbCObKjreEq",
    stateEndpoint: process.env.BotStateEndpoint,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector);

bot.on('conversationUpdate', function (message) {
    if (message.membersAdded) {
        message.membersAdded.forEach(function (identity) {
            if (identity.id == message.address.bot.id) {
                // Bot is joining conversation (page loaded)
                var reply = new builder.Message()
                        .address(message.address)
                        .text("Welcome to DoNotDisclose. Say 'hi' to begin.");
                bot.send(reply);
            } else {
                // User is joining conversation (they sent message)
                // var address = Object.create(message.address);
                // address.user = identity;
                // var reply = new builder.Message()
                //         .address(address)
                //         .text("Hello %s", identity.name);
                // bot.send(reply);
            }
        });
    }
});

bot.dialog('/', [
    function (session, args, next) {
        session.sendTyping();
        setTimeout(function(){
            session.send("Hi! I can help you generate a non-disclosure agreement. Please keep in mind that I am just a bot.");
            session.sendTyping();
        }, 2000);
        setTimeout(function(){
            session.send("Consult an attorney if you need professional assurance that this information is appropriate to your situation.");
            session.sendTyping();
        }, 4000);
        setTimeout(function(){
            session.send("There are different types of NDA's, depending on the circumstances.");
            session.sendTyping();
        }, 6000);

        setTimeout(function(){
            session.send("Your NDA agreement can be a simple, Unilateral agreement to protect only what you disclose or be broader in the form of a “mutual” NDA to protect everyone’s confidential information.");
            session.sendTyping();
        }, 8000);
        
        setTimeout(function(){ 
            builder.Prompts.choice(session, "What type of non-disclosure agreement are you looking to create?", "Unilateral NDA|Multilateral NDA", { maxRetries:0, listStyle: builder.ListStyle.button }); 
        }, 10000);
    },
    function (session, results) { 
        switch (results.response.index) { 
            case 0: 
                session.beginDialog('Unilateral'); 
                break; 
            case 1:
                session.beginDialog('Mutual'); 
                break; 
        };
    }
]);

bot.dialog('Unilateral', [
    function (session, args, next) {
        session.sendTyping();
        setTimeout(function(){
            session.send("Ok, I will help you generate a unilateral non-disclosure agreement.");
            session.sendTyping();
        }, 2000);
        setTimeout(function(){ 
            builder.Prompts.text(session, "What is the full legal name of the individual or company that is disclosing information?");
        }, 4000);
    },
    function (session, results) {
        session.userData.name = results.response;
        builder.Prompts.text(session, "What is the full street address of " + results.response + "?"); 
    },

    function (session, results) {
        session.userData.address = results.response;
        console.log(session.userData.address)
        builder.Prompts.time(session, "What is the effective date of the agreement?"); 
    },
    function (session, results) {
        session.dialogData.time = builder.EntityRecognizer.resolveTime([results.response]);
        session.send("Ok, I am generating the unilateral NDA for " + session.userData.name + " with effective date " + session.dialogData.time.toString(0,14).substring(0,15) + 
                    ". You will receive an email with this document shortly.");


        //Load the docx file as a binary
        var content = fs
            .readFileSync(path.resolve(__dirname, 'input.docx'), 'binary');

        var zip = new JSZip(content);

        var doc = new Docxtemplater();
        doc.loadZip(zip);

        //set the templateVariables
        doc.setData({
            company_name: session.userData.name,
            address: session.userData.address,
            time: session.dialogData.time.toString(0,14).substring(0,15),
            description: 'New Website'
        });

        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            var e = {
                message: error.message,
                name: error.name,
                stack: error.stack,
                properties: error.properties,
            }
            console.log(JSON.stringify({error: e}));
            // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
            throw error;
        }

        var buf = doc.getZip()
                     .generate({type: 'nodebuffer'});

        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);

        // Generate test SMTP service account from ethereal.email
        // Only needed if you don't have a real mail account for testing
        nodemailer.createTestAccount((err, account) => {

            // create reusable transporter object using the default SMTP transport
            let transporter = nodemailer.createTransport({
                host: 'smtp.ethereal.email',
                port: 587,
                secure: false, // true for 465, false for other ports
                auth: {
                    user: account.user, // generated ethereal user
                    pass: account.pass  // generated ethereal password
                }
            });

            // setup email data with unicode symbols
            let mailOptions = {
                from: '"DoNotDisclose" <info@DoNotDisclose.com>', // sender address
                to: 'pietergunst@gmail.com', // list of receivers
                subject: 'Your NDA from DoNotDisclose.com', // Subject line
                text: 'Your NDA from DoNotDisclose.com', // plain text body
                html: '<b>Your NDA is attached to this email</b>', // html body
                // An array of attachments
                attachments: [
                    // File Stream attachment
                    {
                        filename: 'output.docx',
                        path: __dirname + '/output.docx',
                        cid: 'nyan@example.com' // should be as unique as possible
                    }
                ]                
            };

            // send mail with defined transport object
            transporter.sendMail(mailOptions, (error, info) => {
                if (error) {
                    return console.log(error);
                }
                console.log('Message sent: %s', info.messageId);
                // Preview only available when sending through an Ethereal account
                var email_url = nodemailer.getTestMessageUrl(info);
                setTimeout(function(){
                    var actionurl = email_url
                    var card = new builder.ThumbnailCard(session)
                            .title("Download your NDA")
                            .text("I am not able to send emails yet. Click this link to see your NDA.")
                            .images([
                                 builder.CardImage.create(session, "https://s3.amazonaws.com/production.lawgives.com/ep/55/f9/55f9f979e4a99419fb000007.jpeg")
                            ])
                            // .tap(builder.CardAction.openUrl(session, "https://www.legal.io/intake/new?summary=" + session.dialogData.title + "&location=" + location_uri + ""))
                            .buttons([
                                builder.CardAction.openUrl(session, actionurl, "View NDA")
                            ]);
                    var msg = new builder.Message(session).attachments([card]);
                    session.send(msg);
                }, 4000);
                // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@blurdybloop.com>
                // Preview URL: https://ethereal.email/message/WaQKMgKddxQDoou...
            });



        });



    }
]);


bot.dialog('Mutual', [
    function (session, args, next) {
        session.sendTyping();
        setTimeout(function(){
            session.send("Ok, I will help you generate a unilateral NDA.");
            session.sendTyping();
        }, 2000);
    },
    function (session) {
        builder.Prompts.text(session, "What is the full legal name of the company that is disclosing information? (For Example: Do Not Disclose, LLC)");
    },
    function (session, results) {
        session.userData.name = results.response;
        builder.Prompts.number(session, "What is the full street address of " + results.response + "?"); 
    },

    function (session, results) {
        session.userData.address = results.response;
        builder.Prompts.time(session, "What is the effective date of the agreement?"); 
    },
    function (session, results) {
        session.dialogData.time = builder.EntityRecognizer.resolveTime([results.response]);
        session.send("Ok, I am generating the unilateral NDA for " + session.userData.name + 
                    ". You will receive an email with this document shortly.");


        //Load the docx file as a binary
        var content = fs
            .readFileSync(path.resolve(__dirname, 'input.docx'), 'binary');

        var zip = new JSZip(content);

        var doc = new Docxtemplater();
        doc.loadZip(zip);

        //set the templateVariables
        doc.setData({
            company_name: session.userData.name,
            address: session.userData.address,
            time: session.dialogData.time

        });

        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            var e = {
                message: error.message,
                name: error.name,
                stack: error.stack,
                properties: error.properties,
            }
            console.log(JSON.stringify({error: e}));
            // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
            throw error;
        }

        var buf = doc.getZip()
                     .generate({type: 'nodebuffer'});

        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);

        // Generate test SMTP service account from ethereal.email
        // Only needed if you don't have a real mail account for testing
        nodemailer.createTestAccount((err, account) => {

            // create reusable transporter object using the default SMTP transport
            let transporter = nodemailer.createTransport({
                host: 'smtp.ethereal.email',
                port: 587,
                secure: false, // true for 465, false for other ports
                auth: {
                    user: account.user, // generated ethereal user
                    pass: account.pass  // generated ethereal password
                }
            });

            // setup email data with unicode symbols
            let mailOptions = {
                from: '"DoNotDisclose" <info@DoNotDisclose.com>', // sender address
                to: 'pietergunst@gmail.com', // list of receivers
                subject: 'Your NDA from DoNotDisclose.com', // Subject line
                text: 'Your NDA from DoNotDisclose.com', // plain text body
                html: '<b>Your NDA is attached to this email</b>', // html body
                // An array of attachments
                attachments: [
                    // File Stream attachment
                    {
                        filename: 'output.docx',
                        path: __dirname + '/output.docx',
                        cid: 'nyan@example.com' // should be as unique as possible
                    }
                ]                
            };

            // send mail with defined transport object
            transporter.sendMail(mailOptions, (error, info) => {
                if (error) {
                    return console.log(error);
                }
                console.log('Message sent: %s', info.messageId);
                // Preview only available when sending through an Ethereal account
                var email_url = nodemailer.getTestMessageUrl(info);
                setTimeout(function(){
                    var actionurl = email_url
                    var card = new builder.ThumbnailCard(session)
                            .title("Want more options?")
                            .text("Submit a request online and get contacted by attorneys who can assist.")
                            .images([
                                 builder.CardImage.create(session, "https://s3.amazonaws.com/production.lawgives.com/ep/55/f9/55f9f979e4a99419fb000007.jpeg")
                            ])
                            // .tap(builder.CardAction.openUrl(session, "https://www.legal.io/intake/new?summary=" + session.dialogData.title + "&location=" + location_uri + ""))
                            .buttons([
                                builder.CardAction.openUrl(session, actionurl, "Get Started")
                            ]);
                    var msg = new builder.Message(session).attachments([card]);
                    session.send(msg);
                }, 4000);
                // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@blurdybloop.com>
                // Preview URL: https://ethereal.email/message/WaQKMgKddxQDoou...
            });



        });



    }
]);
