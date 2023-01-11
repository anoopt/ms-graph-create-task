import * as core from '@actions/core';
import { PlannerTask } from "@microsoft/microsoft-graph-types";
import { addBusinessDays, format } from 'date-fns';
import Graph from './graph';

async function run() {

    const clientId: string = core.getInput('clientId', { required: true });
    const clientSecret: string = core.getInput('clientSecret', { required: true });
    const tenantId: string = core.getInput('tenantId', { required: true });

    const planId: string = core.getInput('planId', { required: true });
    const title: string = core.getInput('title', { required: true });
    const userId: string = core.getInput('userId', { required: true });
    const bucketId: string = core.getInput('bucketId') ? core.getInput('bucketId') : null;
    const dueBy: string = core.getInput('dueBy');
    const description: string = core.getInput('description');
    const priority: number = core.getInput('priority') ? parseInt(core.getInput('priority')) : 5;
    const orderHint: string = core.getInput('orderHint') ? core.getInput('orderHint') : ' !';

    const graph = new Graph(
        clientId,
        clientSecret,
        tenantId
    );

    const nextWeek: string = format(addBusinessDays(new Date(), 7), 'yyyy-MM-dd');
    const dueDateTime: string = dueBy ? dueBy : `${nextWeek}T10:00:00Z`;

    let assignments: any = {};
    assignments[userId] = {
        "@odata.type": "#microsoft.graph.plannerAssignment",
        orderHint
    };
    
    const task: PlannerTask = {
        planId,
        bucketId,
        title,
        dueDateTime,
        assignments,
        details: {
            description
        },
        priority
    };


    const result: any = await graph.createTask(task);
    core.setOutput('event', result);
}

run();