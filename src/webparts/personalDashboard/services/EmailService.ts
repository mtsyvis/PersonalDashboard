import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory } from '@microsoft/sp-http';
import IEmail from "../models/IEmail";

export default class EmailService {
    public static readonly serviceKey = ServiceKey.create<EmailService>("personal-dashboard:EmailService", EmailService);

    private graph: MSGraphClientFactory;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this.graph = serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }

    public async getEmails(): Promise<IEmail[]> {
        const client = await this.graph.getClient("3");

        try {

            const res = await client
                .api("me/messages")
                .version("v1.0")
                .select("bodyPreview,receivedDateTime,from,subject,webLink")
                .top(5) // todo getting from property pane
                .orderby("receivedDateTime desc")
                .get();

            return res.value || [];
        }
        catch(e) {
            console.warn("Can't read user emails");
            throw e;
        }
    }
}
