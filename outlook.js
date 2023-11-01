const nodemailer = require('nodemailer');

module.exports = sendEmail



function sendEmail() {

    console.log('send email');
    const transporter = nodemailer.createTransport({
        host: 'techvista.onmicrosoft.com',
        service: "Outlook365",
        tls:  { ciphers: 'SSLv3' },
        port: 587,
        secure: false,
        auth: {
            user: 'mohamed.awad@systemsltd.com',
            pass: 'CurlyHoney@66'
        },
      });
      
      const mailOptions = {
        from: 'mohamed.awad@systemsltd.com',
        to: 'mhmdawaddd@gmail.com',
        subject: 'Test Email',
        text: 'This is a test email sent to multiple recipients using Node.js and Nodemailer.'
      };
      
      transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
          console.log(error);
        } else {
          console.log('Email sent: ' + info.response);
        }
      });

}

