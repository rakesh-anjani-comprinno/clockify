import nodemailer from 'nodemailer'
import dotenv from 'dotenv'
dotenv.config()

export const sendEmail = async ({recepentsEmails, subject, body, carbonCopyRecepents = '', attachments = []}) => {
    const transport = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.SMTP_USERNAME,
        pass: process.env.SMTP_PASSWORD,
      }
    })
    const mailOptions =  {
        from: process.env.SMTP_USERNAME,
        to: recepentsEmails,
        cc: carbonCopyRecepents || '',
        subject: subject,
        text: body,
        attachments:attachments
    }

    await transport.sendMail(mailOptions)
    console.log('mail send.....')
}