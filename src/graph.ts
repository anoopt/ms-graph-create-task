import "isomorphic-fetch";
import { AuthProvider, Client, Options } from "@microsoft/microsoft-graph-client";
import { Event, PlannerTask } from "@microsoft/microsoft-graph-types";
import Auth from './auth';
import * as core from '@actions/core';

export default class Graph {

    private auth: Auth;

    constructor(clientId: string, clientSecret: string, tenantId: string) {
        this.auth = new Auth(clientId, clientSecret, tenantId);
    };

    async createTask(task: PlannerTask): Promise<any> {
        const client: Client = await this.getClient();

        if (client) {
            core.info("\u001b[93m⌛ Creating task...");
            try {
                const result: any = await client
                    .api(`/planner/tasks`)
                    .post(task);

                if (result) {
                    core.info("\u001b[32m✅ Task created");
                    console.log(result);
                } else {
                    core.warning("\u001b[33m⚠️ There was an error creating the task");
                }
                return result;
            } catch (error) {
                core.error("\u001b[91m🚨 Error in createTask function.");
                core.error(error);
                core.setFailed(error);
                return null;
            }
        }
        return null;
    };

    private async getClient(): Promise<Client> {
        const accessToken: string = await this.auth.getAccessToken();
        if (accessToken) {
            core.info("\u001b[93m⌛ Getting Graph client...");
            const authProvider: AuthProvider = (done) => {
                done(null, accessToken)
            };
            const options: Options = {
                authProvider
            };
            const client: Client = Client.init(options);
            core.info("\u001b[32m✅ Got Graph client");
            return client;
        }
        return null;
    };
}