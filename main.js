// Connect to email server and retrieve emails as a stream
const Imap = require('imap');

// For sending mail
const nodemailer = require('nodemailer');
//const xoauth2 = require('xoauth2');

// Holds email data collected
let dataHolder = null;

// To parse data into readable format
const {simpleParser} = require('mailparser');
const { parse } = require('path/posix');

// Access to Office 365 server (sender)
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
    to: 'purposeTesting@outlook.com',
    subject: 'Reported',
    text: 'dataHolder'
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
                // All data such as {header, from, to, subject, message ID}
                dataHolder = parsed;
                console.log(dataHolder);
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

//Send attached email out to address
function sendEmail ()
{
    console.log('Testing function');
   transporter.sendMail(mailOptions, function (err, info)
    {
        if (err)
        {
            console.log(err);
            return;
        }
        console.log('Sent: ' + info.response);
    });
};

getEmail();

/*function streamToString (stream) {
    const chunks = [];
    return new Promise((resolve, reject) => {
      stream.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      stream.on('error', (err) => reject(err));
      stream.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    })
  }
  
  //const result = await streamToString(stream)
  streamToString(stream).then(function(response)
  {

  });
  */
