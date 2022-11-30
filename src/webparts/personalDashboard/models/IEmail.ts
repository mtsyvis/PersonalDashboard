export interface IEmail {
    bodyPreview: string;
    from: {
        emailAddress: {
            address: string;
            name: string;
        }
    };
    receivedDateTime: string;
    subject: string;
    webLink: string;
}

export default IEmail;