// Write to files
const fs = require('fs')

// Connect to email server and retrieve emails as a stream
const Imap = require('imap');

// For sending mail
const nodemailer = require('nodemailer');
//const xoauth2 = require('xoauth2');

// To parse data into readable format
const {simpleParser} = require('mailparser');
const { parse } = require('path/posix');

// For converting object
const { stringify } = require('querystring');

// Holds email data collected
let convData = null;

// Access to Office 365 server (obtaining email data)
const imapConfig = 
{
    user: 'testingPurpose1996@outlook.com',
    password: 'Eduardo1996!',
    host: 'outlook.office365.com' ,
    port: 993,
    tls: true
};

// Access to Office 365 server (sender)
let transporter = nodemailer.createTransport(
{
    host: 'smtp-mail.outlook.com',
    //service: 'outlook',
    //host: 'outlook.office365.com',
    port: 587,
    //port: 993,
    //tls: true,
    auth:
    {
        user: 'testingPurpose1996@outlook.com',
        pass: 'Eduardo1996!'
    },
});

let mailOptions = 
{
    from: 'testingPurpose1996@outlook.com',
    to: 'purposeTesting1996@outlook.com',
    subject: 'Reported',
    //text: convData
    attachments:
    {
        filename: 'data.txt'
        //content: convData
        //filename: 'sameple.txt',
        //path: '/Users/dianhernandez/Documents/Office365Extension/sample.txt'
    }
};

// Get Email contents
const getEmail = () => 
{
    try
    {
        const imap = new Imap(imapConfig);
    imap.once('ready', () => 
    {
      imap.openBox('INBOX', false, () => 
      { 
        imap.search(['UNSEEN', ['SINCE', new Date()]], (err, results) => 
        {
          const f = imap.fetch(results, {bodies: ''});
          f.on('message', msg => 
          {
            msg.on('body', stream => 
            {
              simpleParser(stream, async (err, parsed) => 
              {
                // All data such as {header, from, to, subject, message ID} = parsed
                // Parsed object is converted to string = convData
                convData = JSON.stringify(parsed);
                fs.writeFile('data.txt', convData, err => 
                {
                    sendEmail();
                    if (err) 
                    {
                        console.error(err)
                        return
                    }
                })
                
                console.log(parsed);
                // Call to send email with data
                //sendEmail();
                // Save data parsed
            
              });
            });
            msg.once('attributes', attrs => 
            {
              const {uid} = attrs;
              imap.addFlags(uid, ['\\Seen'], () => 
              {
                // Mark the email as read after reading it
              });
            });
          });
          f.once('error', ex => {
            return Promise.reject(ex);
          });
          f.once('end', () => 
          {
            console.log('Done fetching all messages!');
            imap.end();
          });
        });
      });
    });
        imap.once('error', err => 
        {
            console.log(err);
        });

        //Disconnect from server
        imap.once('end', () => {
            console.log('Connection ended.');
        });

        imap.connect();

    }
    // Alert that error occurred while attempting to read
    catch (ex)
    {
        console.log('An error occurred');
    }
};

// Send  email out to address
function sendEmail ()
{
    // Obtains mailOptions specified 
   transporter.sendMail(mailOptions, function (err, info)
    {
        if (err)
        {
            console.log(err);
            return;
        }
    });
};

getEmail();

